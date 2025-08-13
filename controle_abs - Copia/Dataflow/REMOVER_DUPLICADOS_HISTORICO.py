import os
import time
import requests
import gspread
from datetime import datetime
import traceback  # módulo padrão, não adicionar em requirements

# ================== CONFIG CHAT (Webhook via variável de ambiente) ==================
WEBHOOK_URL = os.getenv("WEBHOOK_CHAT_ABS")  # defina no ambiente
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
    chat_text("❌🔥 *Dedup ABS — Erro crítico*\n```\n" + tb_curto + "\n```\n🧯 Tente novamente após alguns minutos.")

# ================== PARÂMETROS VIA AMBIENTE ==================
PLANILHA_ID = os.getenv("PLANILHA_ID")
if not PLANILHA_ID:
    raise RuntimeError("Defina a variável de ambiente PLANILHA_ID (ID da planilha).")

ABA = os.getenv("ABA_HISTORICO", "Historico_agosto")
DRY_RUN = os.getenv("DRY_RUN", "false").strip().lower() in ("1", "true", "yes", "y")

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

# ================== HELPERS ==================
def _last_nonempty_index(values):
    for i in range(len(values) - 1, -1, -1):
        if (values[i] or "").strip() != "":
            return i
    return -1

def _with_backoff(fn, desc="operação", max_tries=6, base_sleep=2):
    """
    Executa fn() com backoff exponencial para erros 429/rate limit.
    """
    sleep_s = base_sleep
    for attempt in range(1, max_tries + 1):
        try:
            return fn()
        except gspread.exceptions.APIError as e:
            s = str(e)
            if "429" in s or "Quota exceeded" in s or "rateLimitExceeded" in s:
                if attempt == max_tries:
                    raise
                print(f"⏳ 429 em {desc}. Tentativa {attempt}/{max_tries}. Aguardando {sleep_s}s... 🐢")
                time.sleep(sleep_s)
                sleep_s = min(sleep_s * 2, 60)
                continue
            raise

# ================== LÓGICA PRINCIPAL ==================
def deduplicar_por_rewrite():
    """
    Dedup EXATO por reescrita:
      - Lê tudo.
      - Mantém primeira ocorrência (A..última coluna do cabeçalho não-vazia).
      - Se houve duplicatas, faz CLEAR() e UPDATE(A1, dados_dedup).
    """
    inicio_str = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    chat_text(
        "🧹✨ *Dedup ABS — Início*\n"
        f"📁 Planilha: `{PLANILHA_ID}`\n"
        f"🗂️ Aba: `{ABA}`\n"
        f"🧪 DRY_RUN: `{DRY_RUN}`\n"
        f"🕒 Início: {inicio_str}"
    )
    try:
        sh = gc.open_by_key(PLANILHA_ID)
        ws = sh.worksheet(ABA)

        todas = _with_backoff(lambda: ws.get_all_values(), "leitura get_all_values")
        if not todas or len(todas) < 2:
            chat_text("ℹ️📝 *Dedup ABS* — Aba vazia ou sem linhas de dados. Nada a remover.")
            return

        cabecalho = todas[0]
        last_idx = _last_nonempty_index(cabecalho)
        if last_idx < 0:
            chat_text("ℹ️🧭 *Dedup ABS* — Cabeçalho sem colunas não vazias. Nada a fazer.")
            return
        ncols = last_idx + 1

        # Dedup mantendo a primeira ocorrência
        vistos = set()
        dedup = [cabecalho[:ncols]]  # começa com o cabeçalho (recortado até ncols)
        removidas = 0

        for row in todas[1:]:
            key = tuple(row[:ncols])  # igualdade EXATA
            if key in vistos:
                removidas += 1
            else:
                vistos.add(key)
                # normaliza o comprimento da linha para ncols
                norm = (row + [""] * (ncols - len(row)))[:ncols]
                dedup.append(norm)

        total_dados = len(todas) - 1

        if removidas == 0:
            chat_text(
                "✅🧼 *Dedup ABS — Concluído*\n"
                f"📊 Linhas de dados: *{total_dados}*\n"
                "🟢 Duplicatas removidas: *0*\n"
                "🧩 Situação: já estava limpo."
            )
            return

        if DRY_RUN:
            chat_text(
                "🔎🧪 *Dedup ABS — DRY RUN*\n"
                f"📊 Linhas de dados (antes): *{total_dados}*\n"
                f"♻️ Duplicatas detectadas: *{removidas}*\n"
                "📝 Ação planejada: *CLEAR + UPDATE* (simulação)."
            )
            return

        # === Reescrita em 2 WRITES (quota-friendly) ===
        _with_backoff(lambda: ws.clear(), "clear()")
        _with_backoff(lambda: ws.update('A1', dedup, value_input_option='RAW'), "update(dedup)")

        finais = len(dedup) - 1
        fim_str = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        chat_text(
            "✅🚀 *Dedup ABS — Concluído*\n"
            f"📉 Linhas de dados (antes): *{total_dados}*\n"
            f"🧽 Duplicatas removidas: *{removidas}*\n"
            f"📈 Linhas de dados (depois): *{finais}*\n"
            f"🏁 Concluído às: {fim_str}"
        )

    except Exception:
        print("❌ Erro durante a execução:")
        print(traceback.format_exc())
        chat_erro()
        raise

if __name__ == "__main__":
    deduplicar_por_rewrite()
