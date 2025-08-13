import os
import requests
import gspread
from datetime import datetime, timedelta
import traceback  # Não adicionar ao requirements.txt

# ========= LOG NO GOOGLE CHAT (via variável de ambiente) =========
WEBHOOK_URL = os.getenv("WEBHOOK_CHAT_Limpeza_padronizacao")
CHAT_TIMEOUT = 10

def _chat_post(payload: dict):
    if not WEBHOOK_URL:
        return
    try:
        r = requests.post(WEBHOOK_URL, json=payload, timeout=CHAT_TIMEOUT)
        if r.status_code >= 300:
            print(f"⚠️ Falha ao enviar Chat (HTTP {r.status_code}): {r.text[:400]}")
    except Exception as ex:
        print(f"⚠️ Exceção ao enviar Chat: {ex}")

def chat_text(msg: str):
    _chat_post({"text": msg})

def chat_erro():
    tb = traceback.format_exc()
    tb_curto = (tb[:1800] + "...") if len(tb) > 1800 else tb
    chat_text("❌ *Limpeza ABS — Erro*\n```\n" + tb_curto + "\n```")

# ========= AUTENTICAÇÃO (Dataflow/Composer) =========
try:
    raw_creds = connections['CONNECTION_IDEA_SP10'].credentials
    scoped_creds = raw_creds.with_scopes([
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive"
    ])
    gc = gspread.authorize(scoped_creds)
except Exception as e:
    print("❌ Erro ao aplicar escopos nas credenciais ou autorizar o acesso ao Google Sheets:", e)
    raise

PLANILHA_ID = '1hSRUlLJkc8iSZc3h7Rdd2tfB4sciamImemyjEySQYBM'
ABA = 'Historico_Gerado_pelo_APP'
BLANK_ROWS_TARGET = 3000  # exatamente 3.000 linhas vazias REAIS ao final

def limpar_e_padronizar_planilha():
    try:
        inicio_str = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        print(f"\n🧹 Iniciando limpeza e padronização - {inicio_str}")
        chat_text(
            f"🧹✨ *Limpeza ABS — Início*\n"
            f"📄 Planilha: `{PLANILHA_ID}`\n"
            f"📑 Aba: `{ABA}`\n"
            f"📏 Linhas vazias alvo: *{BLANK_ROWS_TARGET}*\n"
            f"🕒 Início: {inicio_str}"
        )

        ws = gc.open_by_key(PLANILHA_ID).worksheet(ABA)

        hoje = datetime.now()
        datas_ultimos_3_dias = [
            (hoje - timedelta(days=0)).strftime('%d/%m/%Y'),
            (hoje - timedelta(days=1)).strftime('%d/%m/%Y'),
            (hoje - timedelta(days=2)).strftime('%d/%m/%Y'),
        ]

        # Leitura dos dados
        dados = ws.get_all_values()
        if len(dados) < 2:
            msg = "⚠️ Planilha sem dados suficientes (menos de 2 linhas)."
            print(msg)
            chat_text("ℹ️ *Limpeza ABS* — " + msg)
            return

        header1 = dados[0]
        header2 = dados[1]

        num_colunas = max(
            len(header1),
            len(header2),
            max((len(l) for l in dados[2:]), default=0)
        )

        header1 += [''] * (num_colunas - len(header1))
        header2 += [''] * (num_colunas - len(header2))

        novos_dados = [header1, header2]
        mantidas, removidas = 0, 0

        for linha in dados[2:]:
            linha += [''] * (num_colunas - len(linha))
            valor_x = linha[23] if num_colunas > 23 else ''
            valor_o = linha[14] if num_colunas > 14 else ''

            if not valor_x.strip():
                novos_dados.append(linha)
                mantidas += 1
                continue

            if valor_x == "backup salvo na aba Historico_agosto":
                try:
                    data_linha = datetime.strptime(valor_o, '%d/%m/%Y')
                    data_eh_valida = True
                except Exception:
                    data_eh_valida = False

                if data_eh_valida and (valor_o not in datas_ultimos_3_dias) and (data_linha.date() < hoje.date()):
                    removidas += 1
                    continue

            novos_dados.append(linha)
            mantidas += 1

        antes = len(novos_dados)
        novos_dados = [
            row for i, row in enumerate(novos_dados)
            if i < 2 or any((cell or '').strip() for cell in row)
        ]
        limpas = antes - len(novos_dados)

        # Limpa e reescreve apenas dados
        ws.clear()
        if novos_dados:
            ws.update(range_name='A1', values=novos_dados)

        # Redimensiona para garantir 3.000 linhas reais vazias
        linhas_usadas = len(novos_dados) if novos_dados else 0
        linhas_totais_desejadas = linhas_usadas + BLANK_ROWS_TARGET
        ws.resize(rows=linhas_totais_desejadas, cols=num_colunas)

        fim_str = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        print(f"✅ Processo finalizado. Dados usados: {linhas_usadas} | Linhas vazias reais ao final: {BLANK_ROWS_TARGET}")
        chat_text(
            f"✅🎯 *Limpeza ABS — Concluída*\n"
            f"📊 Linhas mantidas: *{mantidas}*\n"
            f"🗑️ Linhas removidas: *{removidas}*\n"
            f"🧼 Linhas limpas (totalmente vazias): *{limpas}*\n"
            f"📐 Colunas: *{num_colunas}*\n"
            f"✏️ Linhas usadas gravadas: *{linhas_usadas}*\n"
            f"⬜ Linhas vazias reais ao final: *{BLANK_ROWS_TARGET}*\n"
            f"📏 Total de linhas após resize: *{linhas_totais_desejadas}*\n"
            f"⏳ Concluído às: {fim_str}"
        )

    except Exception:
        print("❌ Erro durante a execução da limpeza e padronização:")
        print(traceback.format_exc())
        chat_erro()

if __name__ == "__main__":
    limpar_e_padronizar_planilha()
