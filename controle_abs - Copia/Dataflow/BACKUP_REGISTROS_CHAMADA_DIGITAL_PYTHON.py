import os
import requests
import gspread
from datetime import datetime
import traceback  # módulo padrão, não adicionar em requirements

# ================== CONFIG CHAT (Webhook via variável de ambiente) ==================
WEBHOOK_URL = os.getenv("WEBHOOK_CHAT_ABS")  # defina no ambiente
CHAT_TIMEOUT = 10

def _chat_post(payload: dict):
    """Envia payload para o Google Chat. Se a var de ambiente não existir, não faz nada."""
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
    chat_text("❌ *Backup ABS — Erro*\n```\n" + tb_curto + "\n```")

# ================== CREDENCIAIS (Composer/Dataflow Connection) ==================
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

# ================== PARÂMETROS ==================
ID_ORIGEM   = '1-GRyDj6BUBjRnO2QqMmihxVCZxw3JvJLIrhFHSmgpbI'
ID_DESTINO  = '140f04559QKqzqpOWHwcN5NZhwf38IHH6nSfJuWuriCY'
ABA_ORIGEM  = 'Historico_agosto'
ABA_DESTINO = 'Historico_agosto'

# ================== FUNÇÃO PRINCIPAL ==================
def copiar_substituir_tudo():
    """
    Copia TODOS os valores da aba de origem e substitui COMPLETAMENTE a aba de destino.
    - Sem filtros/validações (ignora Coluna X).
    - Sem marcações de 'backup' na origem.
    - Mantém formatações/validações existentes no destino (clear() remove apenas valores).
    """
    try:
        inicio_str = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        print("\n📌 Iniciando execução (substituição total)...")
        print(f"⏰ Início: {inicio_str}")
        chat_text(
            "🚀 *Backup ABS — Substituição total (COPY -> REPLACE)*\n"
            f"• Origem: `{ID_ORIGEM}` / `{ABA_ORIGEM}`\n"
            f"• Destino: `{ID_DESTINO}` / `{ABA_DESTINO}`\n"
            f"• Início: {inicio_str}"
        )

        # Acessa planilhas/abas
        sh_origem  = gc.open_by_key(ID_ORIGEM)
        sh_destino = gc.open_by_key(ID_DESTINO)
        aba_origem  = sh_origem.worksheet(ABA_ORIGEM)
        aba_destino = sh_destino.worksheet(ABA_DESTINO)

        # Lê TODOS os valores da origem (todas as linhas/colunas preenchidas)
        valores_origem = aba_origem.get_all_values()
        total_linhas = len(valores_origem)
        total_colunas = max((len(l) for l in valores_origem), default=0)

        if total_linhas == 0 or total_colunas == 0:
            msg = "ℹ️ *Backup ABS* — A aba de origem está vazia. Nada para substituir."
            print(msg)
            chat_text(msg)
            return

        print(f"🔎 Origem: {total_linhas} linha(s), {total_colunas} coluna(s) detectadas.")

        # LIMPA a aba de destino (remove os VALORES, preserva validações/formatação)
        print("🧹 Limpando conteúdos da aba de destino...")
        aba_destino.clear()

        # Opcional: garantir que o destino tenha linhas suficientes (normalmente update já expande)
        linhas_atuais_dest = aba_destino.row_count
        if linhas_atuais_dest < total_linhas:
            print(f"➕ Expandindo linhas do destino para {total_linhas}...")
            aba_destino.add_rows(total_linhas - linhas_atuais_dest)

        # Escreve tudo começando em A1
        print(f"⬇️ Gravando {total_linhas} linha(s) no destino (A1)...")
        aba_destino.update('A1', valores_origem, value_input_option='RAW')

        fim_str = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        print(f"✅ Substituição concluída às {fim_str}")
        chat_text(
            "✅ *Backup ABS — Substituição concluída*\n"
            f"• Linhas gravadas: *{total_linhas}*\n"
            f"• Colunas gravadas: *{total_colunas}*\n"
            f"• Concluído às: {fim_str}"
        )

    except Exception:
        print("❌ Erro durante a execução:")
        print(traceback.format_exc())
        chat_erro()
        raise

if __name__ == "__main__":
    copiar_substituir_tudo()
