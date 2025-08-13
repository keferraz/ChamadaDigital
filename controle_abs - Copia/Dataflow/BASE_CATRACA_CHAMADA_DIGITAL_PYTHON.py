import os
import gspread
import numpy as np
from google.cloud import bigquery

SPREADSHEET_ID = os.environ.get("SPREADSHEET_ID")
SHEET_NAME = "dados_catraca_v2"
HEADER = [
    "IDGROOT", "LDAP", "DIA", "ENTRADA", "SAIDA", "DATA",
    "CAD", "Turno_HC", "Data_Turno", "Area_Macro", "Nome"
]

def get_credentials():
    raw_creds = connections['CONNECTION_IDEA_SP10'].credentials
    sheets_scopes = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive",
        "https://www.googleapis.com/auth/cloud-platform"
    ]
    sheets_credentials = raw_creds.with_scopes(sheets_scopes)
    return sheets_credentials

def main():
    print("🚀 Iniciando execução do script de exportação para Google Sheets...")

    if not SPREADSHEET_ID:
        raise ValueError("❌ A variável de ambiente 'SPREADSHEET_ID' não está definida.")
    print(f"📗 Planilha destino: {SPREADSHEET_ID}, aba: {SHEET_NAME}")

    # QUERY BQ
    query = """
    SELECT 
      T.EMPLOYEE_ID AS IDGROOT,
      K.LDAP_USER AS LDAP,
      TIMESTAMP_ADD(T.WORK_START_DATE, INTERVAL 1 HOUR) AS DIA,
      TIMESTAMP_ADD(T.EFFECTIVE_WORK_DAY.START_TIME, INTERVAL 1 HOUR) AS ENTRADA,
      TIMESTAMP_ADD(T.EFFECTIVE_WORK_DAY.END_TIME, INTERVAL 1 HOUR) AS SAIDA,
      DATE(TIMESTAMP_ADD(T.EFFECTIVE_WORK_DAY.START_TIME, INTERVAL 1 HOUR)) AS DATA,
      'BRSP10' AS CAD,
      HC.TURNO AS Turno_HC,
      CASE 
        WHEN HC.TURNO = '3º TURNO' AND (
            EXTRACT(HOUR FROM TIMESTAMP_ADD(T.EFFECTIVE_WORK_DAY.START_TIME, INTERVAL 1 HOUR)) < 7
        )
          THEN DATE_SUB(DATE(TIMESTAMP_ADD(T.EFFECTIVE_WORK_DAY.START_TIME, INTERVAL 1 HOUR)), INTERVAL 1 DAY)
        WHEN HC.TURNO = '5º TURNO' AND (
            EXTRACT(HOUR FROM TIMESTAMP_ADD(T.EFFECTIVE_WORK_DAY.START_TIME, INTERVAL 1 HOUR)) < 6
        )
          THEN DATE_SUB(DATE(TIMESTAMP_ADD(T.EFFECTIVE_WORK_DAY.START_TIME, INTERVAL 1 HOUR)), INTERVAL 1 DAY)
        ELSE DATE(TIMESTAMP_ADD(T.EFFECTIVE_WORK_DAY.START_TIME, INTERVAL 1 HOUR))
      END AS Data_Turno,
      HC.Area_Macro,
      HC.Nome
    FROM `meli-bi-data.WHOWNER.BT_SHP_TYA_EMPLOYEE_TIMECARD` AS T
    JOIN `meli-bi-data.WHOWNER.LK_KRAKEN_USERS` AS K 
      ON T.EMPLOYEE_ID = K.ID
    JOIN (
      SELECT ENTITY_KEY 
      FROM `meli-bi-data.WHOWNER.LK_KRAKEN_ENTITY_ATTRIBUTES`
      WHERE ATTRIBUTE_KEY = 'warehouse' 
        AND DEFAULT_VALUE = 'BRSP10'
    ) AS A 
      ON CAST(T.EMPLOYEE_ID AS STRING) = A.ENTITY_KEY
    LEFT JOIN `meli-sbox.TRANSFORMERS.HC_LAYOUT_IDEA_SP10` AS HC
      ON K.LDAP_USER = HC.LDAP
    WHERE 
      T.EFFECTIVE_WORK_DAY.START_TIME = (
        SELECT MAX(EFFECTIVE_WORK_DAY.START_TIME)
        FROM `meli-bi-data.WHOWNER.BT_SHP_TYA_EMPLOYEE_TIMECARD`
        WHERE EMPLOYEE_ID = T.EMPLOYEE_ID
      )
      AND DATE(T.WORK_START_DATE) >= DATE('2025-06-01')
    ORDER BY ENTRADA DESC
    """

    print("🔑 Conectando ao BigQuery...")
    creds = get_credentials()
    bq_client = bigquery.Client(credentials=creds)
    print("🔎 Executando query no BigQuery...")
    query_job = bq_client.query(query)
    df = query_job.to_dataframe()
    print(f"✅ Query executada! Número de linhas retornadas: {len(df)}")

    df = df[HEADER]
    print(f"🗂️  Colunas do DataFrame: {df.columns.tolist()}")

    # ========= TRATAMENTO DE TIPOS PARA EXPORTAÇÃO GOOGLE SHEETS =========
    for col in df.columns:
        if "db_dtypes" in str(type(df[col].dtype)):
            df[col] = df[col].astype(str)

    for col in df.columns:
        if "Int" in str(df[col].dtype) or "boolean" in str(df[col].dtype):
            df[col] = df[col].astype(str)
        elif np.issubdtype(df[col].dtype, np.datetime64):
            df[col] = df[col].dt.strftime('%Y-%m-%d %H:%M:%S')
        elif str(df[col].dtype).startswith('date'):
            df[col] = df[col].astype(str)
        df[col] = df[col].replace([np.inf, -np.inf], np.nan)
        df[col] = df[col].replace({np.nan: ""})

    print("📡 Conectando à planilha do Google Sheets...")
    gc = gspread.authorize(creds)
    sh = gc.open_by_key(SPREADSHEET_ID)
    worksheet = sh.worksheet(SHEET_NAME)
    print("🔗 Conexão realizada.")

    # Limpa o conteúdo abaixo do cabeçalho, mas mantém as linhas
    num_cols = len(HEADER)
    last_row = worksheet.row_count
    if last_row > 1:
        clear_range = f"A2:{gspread.utils.rowcol_to_a1(last_row, num_cols)}"
        worksheet.batch_clear([clear_range])
        print(f"🧹 Conteúdo antigo limpo no range {clear_range} (mantendo cabeçalho e estrutura).")
    else:
        print("ℹ️ Não há linhas para limpar abaixo do cabeçalho.")

    num_data_rows = len(df)
    total_rows_needed = num_data_rows + 1  # +1 por causa do cabeçalho

    print(f"🔢 Linhas de dados a gravar: {num_data_rows}")
    print(f"📏 Linhas atualmente na aba: {worksheet.row_count}")

    if worksheet.row_count < total_rows_needed:
        rows_to_add = total_rows_needed - worksheet.row_count
        worksheet.add_rows(rows_to_add)
        print(f"➕ {rows_to_add} linhas adicionadas na sheet para comportar os dados.")
    else:
        print("✔️ Quantidade de linhas suficiente na sheet.")

    worksheet.update('A1', [HEADER])
    print("📝 Cabeçalho garantido.")

    if num_data_rows > 0:
        data_range = f"A2:{gspread.utils.rowcol_to_a1(num_data_rows + 1, num_cols)}"
        valores = df.values.tolist()
        print(f"✏️ Gravando {len(valores)} linhas a partir de {data_range} ...")
        worksheet.update(data_range, valores)
        print("✅ Dados inseridos com sucesso!")
    else:
        print("ℹ️ Nenhum dado para inserir.")

    print("🏁 Carga concluída com sucesso!")

if __name__ == "__main__":
    main()
