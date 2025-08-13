import os  # [CHAT]
import requests  # [CHAT]
import gspread
from datetime import datetime
import traceback  # Módulo padrão, NÃO incluir no requirements.txt

# ============== CHAT (Webhook via variável de ambiente) ==============
WEBHOOK_URL = os.getenv("WEBHOOK_CHAT_ABS")  # defina no ambiente
CHAT_TIMEOUT = 10

def _chat_post(payload: dict):
    """Envia payload para Google Chat. Se a var de ambiente não existir, ignora silenciosamente."""
    if not WEBHOOK_URL:
        return
    try:
        r = requests.post(WEBHOOK_URL, json=payload, timeout=CHAT_TIMEOUT)
        if r.status_code >= 300:
            print(f"⚠️ Falha ao enviar Chat (HTTP {r.status_code}): {r.text[:400]}")
    except Exception as ex:
        print(f"⚠️ Exceção ao enviar Chat: {ex}")

def chat_text(msg: str):
    """Mensagem de texto simples (mais robusta)."""
    _chat_post({"text": msg})

def chat_erro():
    """Mensagem de erro com traceback truncado."""
    tb = traceback.format_exc()
    tb_curto = (tb[:1800] + "...") if len(tb) > 1800 else tb
    chat_text("❌ *Backup ABS — Erro*\n```\n" + tb_curto + "\n```")

# ============== CREDENCIAIS (Composer/Dataflow Connection) ==============
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

ID_ORIGEM = '1hSRUlLJkc8iSZc3h7Rdd2tfB4sciamImemyjEySQYBM'
ID_DESTINO = '1-GRyDj6BUBjRnO2QqMmihxVCZxw3JvJLIrhFHSmgpbI'
ABA_ORIGEM = 'Historico_Gerado_pelo_APP'
ABA_DESTINO = 'Historico_agosto'
COLUNA_X = 24  # Coluna X = 24 (índice 23)
COLUNA_B = 2   # Coluna B = 2 (índice 1)

def copiar_e_marcar_beckup():
    try:
        inicio_str = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        print("\n📌 Iniciando execução...")
        print(f"⏰ Início: {inicio_str}")
        # [CHAT]
        chat_text(
            "🚀 *Backup ABS — Início*\n"
            f"• Origem: `{ID_ORIGEM}` / `{ABA_ORIGEM}`\n"
            f"• Destino: `{ID_DESTINO}` / `{ABA_DESTINO}`\n"
            f"• Início: {inicio_str}"
        )

        aba_origem = gc.open_by_key(ID_ORIGEM).worksheet(ABA_ORIGEM)
        aba_destino = gc.open_by_key(ID_DESTINO).worksheet(ABA_DESTINO)

        # Ler todos os dados da origem (dados começam na linha 3)
        valores_origem = aba_origem.get_all_values()
        if not valores_origem or len(valores_origem) < 3:
            print("⚠️ Nenhum dado para copiar.")
            chat_text("ℹ️ *Backup ABS* — Nenhum dado elegível para copiar.")  # [CHAT]
            return

        cabecalho = valores_origem[1]      # Cabeçalho está na linha 2 (índice 1)
        dados_origem = valores_origem[2:]  # Dados a partir da linha 3 (índice 2)

        # Seleciona apenas os registros com coluna X vazia E coluna B preenchida
        linhas_para_backup = []
        indices_para_marcar = []
        for idx, linha in enumerate(dados_origem):
            coluna_x_vazia = (len(linha) < COLUNA_X or linha[COLUNA_X-1].strip() == '')
            coluna_b_preenchida = (len(linha) > 1 and linha[1].strip() != '')
            if coluna_x_vazia and coluna_b_preenchida:
                linhas_para_backup.append(linha)
                indices_para_marcar.append(idx+3)  # +3 porque dados começam na linha 3 da planilha

        if not linhas_para_backup:
            print("Nenhum dado novo para backup.")
            chat_text("ℹ️ *Backup ABS* — Nenhum dado novo para backup.")  # [CHAT]
            return

        # Checar última linha preenchida no destino
        valores_destino = aba_destino.get_all_values()
        if not valores_destino:
            print("Planilha de destino vazia, adicionando cabeçalho.")
            aba_destino.append_row(cabecalho, value_input_option='RAW')
            ultima_linha = 1
        else:
            ultima_linha = len(valores_destino)

        # Expande linhas do destino se necessário
        linhas_necessarias = ultima_linha + len(linhas_para_backup)
        if aba_destino.row_count < linhas_necessarias:
            print(f"Expandindo a planilha destino para {linhas_necessarias} linhas...")
            aba_destino.add_rows(linhas_necessarias - aba_destino.row_count)
            chat_text(f"📈 *Backup ABS* — Destino expandido para {linhas_necessarias} linhas.")  # [CHAT]

        # Inserção em bloco no destino
        primeira_nova_linha = ultima_linha + 1
        print(f"Inserindo {len(linhas_para_backup)} linhas no destino, começando pela linha {primeira_nova_linha}")
        aba_destino.update(f'A{primeira_nova_linha}', linhas_para_backup)
        chat_text(  # [CHAT]
            f"⬇️ *Backup ABS* — Inserindo *{len(linhas_para_backup)}* linhas no destino a partir da linha *{primeira_nova_linha}*."
        )

        # Marcar "backup salvo na aba {ABA_DESTINO}" na coluna X da aba origem
        # Se os índices não são sequenciais, faça individualmente para evitar sobrescrever linhas erradas!
        for idx in indices_para_marcar:
            aba_origem.update(f'X{idx}', [[f"backup salvo na aba {ABA_DESTINO}"]])

        fim_str = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        print(f"✅ {len(linhas_para_backup)} linhas copiadas e marcadas na origem.")
        print(f"🏁 Execução concluída às {fim_str}")
        chat_text(  # [CHAT]
            "✅ *Backup ABS — Execução concluída*\n"
            f"• Linhas copiadas: *{len(linhas_para_backup)}*\n"
            f"• Primeira nova linha no destino: *{primeira_nova_linha}*\n"
            f"• Concluído às: {fim_str}"
        )

    except Exception:
        print("❌ Erro durante a execução:")
        print(traceback.format_exc())
        chat_erro()  # [CHAT]

if __name__ == "__main__":
    copiar_e_marcar_beckup()
