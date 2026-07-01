"""
Configuração simples de usuários - Fase Multiusuário Premium.

Perfis disponíveis:
- admin: acesso total. Pode importar Excel, disparar alertas de jornada e ver auditoria.
- consulta: acesso de leitura. Pode consultar dashboards, alertas, extratos e baixar Excel.

IMPORTANTE:
- Troque as senhas padrão antes de liberar para a equipe.
- Para adicionar usuário, copie um bloco do dicionário USUARIOS.
- O campo "perfil" deve ser "admin" ou "consulta".
"""

import os

USUARIOS = {
    "admin": {
        "senha": os.environ.get("TEMPO_FECHADO_ADMIN_PASSWORD") or "admin123",
        "nome": "Administrador",
        "perfil": "admin",
    },
    "max": {
        "senha": os.environ.get("TEMPO_FECHADO_MAX_PASSWORD") or "ponto123",
        "nome": "Max",
        "perfil": "admin",
    },
    "consulta": {
        "senha": os.environ.get("TEMPO_FECHADO_CONSULTA_PASSWORD") or "consulta123",
        "nome": "Usuário Consulta",
        "perfil": "consulta",
    },
}
