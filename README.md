# Tempo Fechado v8.21.38

Sistema web para upload manual de PDFs de ponto eletrônico, processamento de marcações e geração de relatórios.

## Requisitos

- Windows 10/11 (64 bits)
- Python 3.9+
- Opcional: [Inno Setup Portable](https://jrsoftware.org/isdl.php) para gerar instalador

## Instalação rápida

```batch
pip install flask pandas openpyxl numpy python-dateutil pytz werkzeug jinja2 pywin32 pdfplumber pdfminer.six waitress
```

## Configuração

Copie o arquivo `.env.example` como `.env` e ajuste as variáveis:

```env
# Chave secreta do Flask (troque por um valor aleatório)
TEMPO_FECHADO_SECRET_KEY=bbf4b37477c714f4df81447fcdfdc58e1be480cc0342885f

# Senhas dos usuários padrão
TEMPO_FECHADO_ADMIN_PASSWORD=admin123
TEMPO_FECHADO_MAX_PASSWORD=ponto123
TEMPO_FECHADO_CONSULTA_PASSWORD=consulta123

# Caminho raiz dos PDFs
TEMPO_FECHADO_PASTA_RAIZ=%USERPROFILE%\ponto_pdfs

# E-mail do remetente filtro (Outlook)
TEMPO_FECHADO_REMETENTE_FILTRO=seu.email@empresa.com
```

> **Importante:** Troque a `TEMPO_FECHADO_SECRET_KEY` por um valor aleatório em produção. Altere também as senhas padrão antes de liberar para a equipe.

## Executar

### Modo teste (servidor Flask embutido)

Duplo clique em `iniciar_modo_teste.bat` ou:

```batch
python robo_ponto_web.py
```

Acesse: http://127.0.0.1:5050

### Modo produção (waitress)

Duplo clique em `iniciar_servidor_rede.bat` ou:

```batch
python -m waitress --host=0.0.0.0 --port=5050 robo_ponto_web:app
```

Acesse de qualquer máquina da rede: http://<IP_DO_SERVIDOR>:5050

### Launcher completo (recomendado)

```batch
python TempoFechado_Launcher.py
```

Inicia o servidor e abre o navegador automaticamente.

## Usuários padrão

| Usuário | Senha | Perfil |
|---------|-------|--------|
| `admin` | definida no `.env` | admin |
| `max` | definida no `.env` | admin |
| `consulta` | definida no `.env` | consulta |

> Altere as senhas no `.env` ou via interface de Administração após o primeiro login.

## Estrutura de pastas

```
.
├── robo_ponto_web.py          # Aplicação Flask principal
├── TempoFechado_Launcher.py   # Launcher que inicia o servidor e abre o navegador
├── user_db.py                 # Camada de banco de usuários (SQLite)
├── usuarios_config.py         # Configuração legada de usuários
├── alertas_notificacao_config.py  # Config de e-mail/WhatsApp
├── configuracoes_robo.modelo.json  # Modelo de configuração do robô
├── Script_Launcher_v5_filtrado.py  # Download de anexos do Outlook
├── Script_Robo_Ponto_v7_PRO.py     # Motor de processamento de PDFs/Excel
├── templates/                 # Templates HTML (index, login)
├── static/                    # Assets estáticos (CSS, JS, imagens)
├── tools/                     # Scripts auxiliares
├── assets/                    # Ícones do aplicativo
├── .env                       # Variáveis de ambiente (NÃO versionar)
└── .gitignore
```

## Gerar EXE (PyInstaller)

```batch
GERAR_TempoFechado_EXE.bat
```

Gera `dist\TempoFechado\TempoFechado.exe`.

## Gerar instalador (Inno Setup)

1. Execute `GERAR_TempoFechado_EXE.bat` primeiro
2. Execute `GERAR_Instalador_TempoFechado.bat`

Gera `installer_output\Setup_TempoFechado_v8_21_38.exe`.

O instalador cria o aplicativo em `%LOCALAPPDATA%\Tempo Fechado` (sem necessidade de administrador).

## Dados do usuário

Os dados ficam em `%USERPROFILE%\ponto_pdfs\`:

```
%USERPROFILE%\ponto_pdfs\
├── entrada\          # PDFs enviados pelo usuário
├── saida\            # Relatórios Excel gerados
├── processados\      # PDFs já processados
├── backup\           # Backups automáticos
├── erro\             # PDFs com erro
├── logs\             # Logs do sistema
├── config\           # Configurações do robô
└── controle\         # Arquivos de controle do robô
```
