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

USUARIOS = {
    "admin": {
        "senha": "admin123",
        "nome": "Administrador",
        "perfil": "admin",
    },
    "max": {
        "senha": "ponto123",
        "nome": "Max",
        "perfil": "admin",
    },
    "consulta": {
        "senha": "consulta123",
        "nome": "Usuário Consulta",
        "perfil": "consulta",
    },
    # Exemplo para liberar um novo usuário somente consulta:
    # "maria": {
    #     "senha": "trocar123",
    #     "nome": "Maria Silva",
    #     "perfil": "consulta",
    # },
}
