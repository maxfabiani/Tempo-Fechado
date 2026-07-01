# ============================================================
# CONFIGURAÇÃO - ALERTAS DE JORNADA POR E-MAIL / WHATSAPP
# ============================================================
# 1) Para ativar e-mail: preencha SMTP_* e deixe EMAIL_ATIVO = True.
# 2) Para ativar WhatsApp: use um provedor com API HTTP, como Z-API, Twilio,
#    UltraMsg ou outro gateway corporativo, e preencha WHATSAPP_*.
# 3) O sistema evita reenvio da mesma ocorrência usando data/alertas_jornada_enviados.json.

EMAIL_ATIVO = False
SMTP_HOST = "smtp.office365.com"   # Ex.: smtp.office365.com, smtp.gmail.com
SMTP_PORT = 587
SMTP_USAR_TLS = True
SMTP_USUARIO = ""                  # Ex.: seu.email@empresa.com.br
SMTP_SENHA = ""                    # Senha/app password/token SMTP
EMAIL_REMETENTE = ""               # Ex.: seu.email@empresa.com.br
EMAIL_DESTINATARIOS = [
    # "gestor@empresa.com.br",
]

WHATSAPP_ATIVO = False
WHATSAPP_API_URL = ""              # Endpoint do provedor WhatsApp
WHATSAPP_TOKEN = ""                # Token/credencial do provedor
WHATSAPP_DESTINATARIOS = [
    # "5521999999999",
]

# Comportamento
ENVIAR_AO_IMPORTAR_EXCEL = True
ENVIAR_APENAS_OCORRENCIAS_NOVAS = True
LIMITE_OCORRENCIAS_POR_ENVIO = 50
