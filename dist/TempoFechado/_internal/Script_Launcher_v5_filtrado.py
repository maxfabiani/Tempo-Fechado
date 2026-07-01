import json
import os
import re
import logging
import subprocess
import sys
from logging.handlers import RotatingFileHandler
from pathlib import Path

import win32com.client

# ================= CONFIG =================
BASE_DIR = Path.home() / "ponto_pdfs"
PASTA_ENTRADA = BASE_DIR / "entrada"
PASTA_CONTROLE = BASE_DIR / "controle"

ARQUIVO_LOG = PASTA_CONTROLE / "download_anexos.log"
LOCK_FILE = PASTA_CONTROLE / "download.lock"
ARQUIVO_RESULTADO_ROBO = PASTA_CONTROLE / "robo_ponto_resultado.json"

APP_DIR = Path(__file__).resolve().parent
SCRIPT_ROBO = APP_DIR / "Script_Robo_Ponto_v7_PRO.py"
ARQUIVO_CONFIG_ROBO_LOCAL_V81743 = APP_DIR / "data" / "configuracoes_robo.json"
ARQUIVO_CONFIG_ROBO_USUARIO_V81743 = BASE_DIR / "config" / "configuracoes_robo.json"

ASSUNTO_FILTRO = "Análise de controle de marcações"
ASSUNTOS_FILTRO_EXTRA = [
    "analise de controle de marcacoes",
    "análise de controle de marcações",
    "controle de marcacoes",
    "controle de marcações",
]
ACEITAR_PDF_DA_FERNANDA_INDEPENDENTE_ASSUNTO = False
REMETENTE_FILTRO = "fernanda.josino@engemon.com.br"

DIAGNOSTICO_OUTLOOK_V81723 = True
ANALISAR_TAMBEM_LIDOS_RECENTES = False
LIMITE_MENSAGENS_DIAGNOSTICO_PADRAO = 500
LIMITE_MENSAGENS_DIAGNOSTICO_MIN = 50
LIMITE_MENSAGENS_DIAGNOSTICO_MAX = 5000


def carregar_config_robo_v81743() -> dict:
    cfg = {
        "capturar_emails_lidos": False,
        "limite_emails_outlook": LIMITE_MENSAGENS_DIAGNOSTICO_PADRAO,
        "remetente_filtro": REMETENTE_FILTRO,
        "assunto_filtro": ASSUNTO_FILTRO,
    }
    for arquivo_cfg in (ARQUIVO_CONFIG_ROBO_LOCAL_V81743, ARQUIVO_CONFIG_ROBO_USUARIO_V81743):
        try:
            if arquivo_cfg.exists():
                with open(arquivo_cfg, "r", encoding="utf-8") as f:
                    dados = json.load(f) or {}
                if isinstance(dados, dict):
                    cfg.update({k: v for k, v in dados.items() if v is not None})
        except Exception as e:
            try:
                print(f"[TempoFechado] Falha ao ler configuracoes_robo.json ({arquivo_cfg}): {e}")
            except Exception:
                pass
    return cfg


CONFIG_ROBO_V81743 = carregar_config_robo_v81743()
CAPTURAR_EMAILS_LIDOS_V81743 = bool(CONFIG_ROBO_V81743.get("capturar_emails_lidos", False))
REMETENTE_FILTRO = str(CONFIG_ROBO_V81743.get("remetente_filtro") or REMETENTE_FILTRO).strip() or REMETENTE_FILTRO
ASSUNTO_FILTRO = str(CONFIG_ROBO_V81743.get("assunto_filtro") or ASSUNTO_FILTRO).strip() or ASSUNTO_FILTRO

def normalizar_limite_emails_outlook_v8216(valor) -> int:
    try:
        limite = int(valor or LIMITE_MENSAGENS_DIAGNOSTICO_PADRAO)
    except Exception:
        limite = LIMITE_MENSAGENS_DIAGNOSTICO_PADRAO
    limite = max(LIMITE_MENSAGENS_DIAGNOSTICO_MIN, min(limite, LIMITE_MENSAGENS_DIAGNOSTICO_MAX))
    return limite

LIMITE_MENSAGENS_DIAGNOSTICO = normalizar_limite_emails_outlook_v8216(
    CONFIG_ROBO_V81743.get("limite_emails_outlook", LIMITE_MENSAGENS_DIAGNOSTICO_PADRAO)
)


def python_para_subprocesso_sem_console():
    try:
        exe = Path(sys.executable)
        if sys.platform.startswith("win") and exe.name.lower() == "python.exe":
            pythonw = exe.with_name("pythonw.exe")
            if pythonw.exists():
                return str(pythonw)
        return str(exe)
    except Exception:
        return sys.executable


def montar_execucao_robo_v82126():
    env = os.environ.copy()
    env.setdefault("PYTHONUTF8", "1")
    env.setdefault("PYTHONIOENCODING", "utf-8")
    if getattr(sys, "frozen", False):
        env["TEMPO_FECHADO_RUN_SCRIPT"] = str(SCRIPT_ROBO)
        env["TEMPO_FECHADO_RUN_SCRIPT_CWD"] = str(SCRIPT_ROBO.parent)
        env["TEMPO_FECHADO_RUN_SCRIPT_MODE"] = "robo_via_launcher"
        return [str(sys.executable)], env
    return [python_para_subprocesso_sem_console(), str(SCRIPT_ROBO)], env


def kwargs_subprocess_sem_console():
    kwargs = {}
    if sys.platform.startswith("win"):
        try:
            startupinfo = subprocess.STARTUPINFO()
            startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
            startupinfo.wShowWindow = 0
            kwargs["startupinfo"] = startupinfo
            kwargs["creationflags"] = getattr(subprocess, "CREATE_NO_WINDOW", 0)
        except Exception:
            pass
    return kwargs

# ================= SETUP =================
PASTA_ENTRADA.mkdir(parents=True, exist_ok=True)
PASTA_CONTROLE.mkdir(parents=True, exist_ok=True)

logger = logging.getLogger("download")
logger.setLevel(logging.INFO)
logger.handlers.clear()

handler = RotatingFileHandler(
    ARQUIVO_LOG,
    maxBytes=1_000_000,
    backupCount=3,
    encoding="utf-8",
)
formatter = logging.Formatter("[%(asctime)s] %(message)s")
handler.setFormatter(formatter)
logger.addHandler(handler)
logger.addHandler(logging.StreamHandler())


# ================= LOCK =================
def criar_lock():
    if LOCK_FILE.exists():
        logger.warning("Já está em execução.")
        sys.exit(1)
    LOCK_FILE.write_text("rodando", encoding="utf-8")


def remover_lock():
    if LOCK_FILE.exists():
        LOCK_FILE.unlink()


# ================= RESULTADO =================
def ler_resultado_robo():
    if not ARQUIVO_RESULTADO_ROBO.exists():
        return None

    with open(ARQUIVO_RESULTADO_ROBO, "r", encoding="utf-8") as f:
        return json.load(f)


# ================= OUTLOOK =================
def obter_email_remetente(msg) -> str:
    try:
        sender_email = str(msg.SenderEmailAddress).strip()
    except Exception:
        sender_email = ""

    try:
        if getattr(msg, "SenderEmailType", "").upper() == "EX":
            exch_user = msg.Sender.GetExchangeUser()
            if exch_user and exch_user.PrimarySmtpAddress:
                return str(exch_user.PrimarySmtpAddress).strip().lower()
    except Exception:
        pass

    return sender_email.lower()


def salvar_anexo_com_nome_unico(anexo, pasta_destino: Path) -> Path:
    nome_arquivo = str(anexo.FileName).strip()
    caminho_destino = pasta_destino / nome_arquivo

    if caminho_destino.exists():
        stem = caminho_destino.stem
        suffix = caminho_destino.suffix
        contador = 1
        while caminho_destino.exists():
            caminho_destino = pasta_destino / f"{stem}_{contador}{suffix}"
            contador += 1

    anexo.SaveAsFile(str(caminho_destino))
    return caminho_destino




def normalizar_assunto_v81735(valor: str) -> str:
    import unicodedata
    texto = str(valor or "").strip().lower()
    texto = unicodedata.normalize("NFKD", texto)
    texto = "".join(ch for ch in texto if not unicodedata.combining(ch))
    for prefixo in ["re:", "res:", "enc:", "fw:", "fwd:"]:
        while texto.startswith(prefixo):
            texto = texto[len(prefixo):].strip()
    texto = re.sub(r"[^a-z0-9]+", " ", texto)
    texto = re.sub(r"\s+", " ", texto).strip()
    return texto


def assunto_compativel_v81734(assunto: str) -> bool:
    """v8.17.36: aceita somente a família 'Análise de controle de marcação/marcações'.

    A regra exige:
    - controle
    - raiz marcac, que cobre marcacao e marcacoes
    """
    assunto_norm = normalizar_assunto_v81735(assunto)
    filtro_norm = normalizar_assunto_v81735(ASSUNTO_FILTRO)
    if filtro_norm and filtro_norm not in {
        "analise de controle de marcacoes",
        "analise de controle de marcacao",
    }:
        return filtro_norm in assunto_norm
    return "controle" in assunto_norm and "marcac" in assunto_norm

def baixar_anexos_outlook() -> int:
    logger.info("========== DIAGNOSTICO OUTLOOK v8.21.20 - INSTANCIA UNICA ==========")
    logger.info(f"Limite configurado para leitura do Outlook: {LIMITE_MENSAGENS_DIAGNOSTICO} e-mail(s) recentes.")
    logger.info("Conectando ao Outlook...")

    try:
        outlook = win32com.client.Dispatch("Outlook.Application")
        namespace = outlook.GetNamespace("MAPI")
        inbox = namespace.GetDefaultFolder(6)
        logger.info(f"Pasta Outlook usada: {getattr(inbox, 'Name', 'Inbox')}")
    except Exception as e:
        logger.exception(f"Falha ao conectar no Outlook/MAPI: {e}")
        return 0

    try:
        if CAPTURAR_EMAILS_LIDOS_V81743:
            mensagens = inbox.Items
            logger.info("Modo configurado v8.19: capturando PDFs de e-mails nao lidos e tambem de e-mails lidos recentes.")
        elif ANALISAR_TAMBEM_LIDOS_RECENTES:
            mensagens = inbox.Items
            logger.info("Modo diagnostico: listando e-mails recentes; somente nao lidos podem ter PDFs capturados.")
        else:
            mensagens = inbox.Items.Restrict("[Unread] = true")
            logger.info("Modo padrao: analisando e capturando somente e-mails nao lidos.")
        mensagens.Sort("[ReceivedTime]", True)
    except Exception as e:
        logger.exception(f"Falha ao preparar lista de mensagens: {e}")
        return 0

    total_analisados = 0
    total_nao_lidos = 0
    total_remetente_ok = 0
    total_assunto_ok = 0
    total_com_pdf = 0
    total_baixados = 0
    remetentes_vistos = {}
    assuntos_vistos = []
    assuntos_fernanda_v81736 = []

    for msg in mensagens:
        try:
            total_analisados += 1
            if total_analisados > LIMITE_MENSAGENS_DIAGNOSTICO:
                logger.info(f"Limite configurado atingido: {LIMITE_MENSAGENS_DIAGNOSTICO} mensagem(ns). Ajuste este valor em Administração do Robô > Outlook se necessário.")
                break

            try:
                unread = bool(getattr(msg, "Unread", False))
            except Exception:
                unread = False

            if unread:
                total_nao_lidos += 1

            remetente = obter_email_remetente(msg)
            assunto = str(getattr(msg, "Subject", "") or "").strip()
            recebido = str(getattr(msg, "ReceivedTime", "") or "")

            remetentes_vistos[remetente or "(sem remetente)"] = remetentes_vistos.get(remetente or "(sem remetente)", 0) + 1
            if len(assuntos_vistos) < 15:
                assuntos_vistos.append(f"{recebido} | unread={unread} | {remetente} | {assunto}")
            if remetente.lower() == REMETENTE_FILTRO.lower():
                assuntos_fernanda_v81736.append(f"{recebido} | unread={unread} | {assunto}")

            logger.info(f"Analisando #{total_analisados}: unread={unread} | remetente={remetente} | assunto={assunto}")

            # v8.17.43: politica configuravel pela Administracao do Robo.
            # Padrao seguro: baixar PDFs apenas de e-mails NAO LIDOS.
            # Quando habilitado, permite incluir e-mails lidos recentes.
            if not unread and not CAPTURAR_EMAILS_LIDOS_V81743:
                logger.info("  -> ignorado: e-mail ja lido. Politica atual captura somente nao lidos.")
                continue

            if remetente.lower() != REMETENTE_FILTRO.lower():
                logger.info(f"  -> ignorado: remetente diferente do filtro ({REMETENTE_FILTRO}).")
                continue

            total_remetente_ok += 1

            anexos = msg.Attachments
            qtd_anexos = getattr(anexos, "Count", 0)
            logger.info(f"  -> anexos encontrados: {qtd_anexos}")

            tem_pdf = False
            try:
                for j in range(1, anexos.Count + 1):
                    nome_previa = str(anexos.Item(j).FileName).strip()
                    if nome_previa.lower().endswith(".pdf"):
                        tem_pdf = True
                        break
            except Exception as e:
                logger.warning(f"  -> falha ao verificar anexos previamente: {e}")

            if not assunto_compativel_v81734(assunto):
                logger.info("  -> ignorado: assunto fora da família 'Análise de controle de marcação/marcações'.")
                continue

            total_assunto_ok += 1
            logger.info(f"  -> e-mail compativel encontrado pela regra v8.19 | Assunto: {assunto} | Remetente: {remetente}")

            baixou_pdf = False

            for i in range(1, anexos.Count + 1):
                anexo = anexos.Item(i)
                nome_arquivo = str(anexo.FileName).strip()
                logger.info(f"     anexo #{i}: {nome_arquivo}")

                if not nome_arquivo.lower().endswith(".pdf"):
                    logger.info("       -> ignorado: nao e PDF.")
                    continue

                total_com_pdf += 1
                caminho_salvo = salvar_anexo_com_nome_unico(anexo, PASTA_ENTRADA)
                logger.info(f"       -> PDF salvo em: {caminho_salvo}")
                total_baixados += 1
                baixou_pdf = True

            if baixou_pdf:
                try:
                    msg.Unread = False
                    msg.Save()
                    logger.info("  -> e-mail marcado como lido.")
                except Exception as e:
                    logger.warning(f"  -> PDF baixado, mas nao foi possivel marcar como lido: {e}")
            else:
                logger.info("  -> e-mail compativel encontrado, mas sem PDF anexo.")

        except Exception as e:
            logger.exception(f"Erro ao processar e-mail #{total_analisados}: {e}")

    logger.info("========== RESUMO DIAGNOSTICO OUTLOOK ==========")
    logger.info(f"Total de e-mails analisados: {total_analisados}")
    logger.info(f"Total de e-mails nao lidos: {total_nao_lidos}")
    logger.info(f"Total de e-mails da Fernanda: {total_remetente_ok}")
    logger.info(f"Total de e-mails da Fernanda com assunto compativel: {total_assunto_ok}")
    logger.info(f"Total de PDFs compativeis encontrados: {total_com_pdf}")
    logger.info(f"Total de PDFs baixados: {total_baixados}")

    logger.info("Remetentes vistos no diagnostico:")
    for remetente, qtd in sorted(remetentes_vistos.items(), key=lambda x: x[1], reverse=True)[:20]:
        logger.info(f"  {qtd}x | {remetente}")

    logger.info("Amostra de assuntos recentes:")
    for linha in assuntos_vistos:
        logger.info(f"  {linha}")

    logger.info("Todos os assuntos da Fernanda encontrados na janela analisada:")
    if assuntos_fernanda_v81736:
        for linha in assuntos_fernanda_v81736:
            logger.info(f"  {linha}")
    else:
        logger.info("  Nenhum e-mail da Fernanda encontrado na janela analisada.")

    if total_baixados == 0:
        logger.warning("Nenhum PDF baixado. Verifique Outlook aberto/perfil ativo, e-mail lido, remetente, assunto, anexos e subpastas.")

    return total_baixados


# ================= MAIN =================
def main():
    criar_lock()
    try:
        politica_outlook = "inclui lidos recentes" if CAPTURAR_EMAILS_LIDOS_V81743 else "somente nao lidos"
        logger.info(f"Iniciando execução do robô completo - v8.21.20 | Outlook configuravel ({politica_outlook}) | limite={LIMITE_MENSAGENS_DIAGNOSTICO}")

        total_baixados = baixar_anexos_outlook()

        if total_baixados == 0:
            pdfs_existentes = list(PASTA_ENTRADA.glob("*.pdf"))
            if pdfs_existentes:
                logger.warning(f"Nenhum PDF novo foi baixado do Outlook, mas ha {len(pdfs_existentes)} PDF(s) na pasta entrada. O robo sera executado mesmo assim.")
            else:
                logger.error("Nenhum PDF foi baixado do Outlook e a pasta entrada esta vazia. Robo sera encerrado antes do processamento.")
                return 1

        logger.info("Executando robô de ponto...")
        comando_robo, env_robo = montar_execucao_robo_v82126()
        resultado = subprocess.run(
            comando_robo,
            capture_output=True,
            text=True,
            env=env_robo,
            **kwargs_subprocess_sem_console(),
        )

        logger.info(f"Return code: {resultado.returncode}")

        if resultado.stdout:
            logger.info(resultado.stdout)

        if resultado.stderr:
            logger.error(resultado.stderr)

        resultado_json = ler_resultado_robo()

        if resultado.returncode != 0:
            logger.error("Robô retornou erro.")
            return 1

        if not resultado_json:
            logger.error("Robô não gerou JSON.")
            return 1

        if resultado_json.get("status") != "sucesso":
            logger.warning(f"Execução parcial: {resultado_json}")
            return 1

        logger.info(f"Sucesso total: {resultado_json}")
        return 0

    finally:
        remover_lock()


if __name__ == "__main__":
    sys.exit(main())
