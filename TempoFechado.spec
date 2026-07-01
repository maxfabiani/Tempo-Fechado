from pathlib import Path
from PyInstaller.utils.hooks import collect_submodules
block_cipher=None
ROOT=Path.cwd()
datas=[("templates","templates"),("static","static")]
if (ROOT/"assets").exists(): datas.append(("assets","assets"))
for nome in ["users.db","usuarios.db","tempo_fechado.db","config.json","config_robo.json","alertas_notificacao_config.json"]:
    p=ROOT/nome
    if p.exists(): datas.append((str(p), "."))
for nome in ["Script_Launcher_v5_filtrado.py", "Script_Robo_Ponto_v7_PRO.py"]:
    p = ROOT / nome
    if p.exists():
        datas.append((str(p), "."))
if (ROOT/"tools").exists(): datas.append(("tools","tools"))
hiddenimports=[
    "flask","werkzeug","jinja2","sqlite3","pandas","numpy","openpyxl","dateutil","pytz",
    "logging.handlers","csv","dataclasses","typing","email.message",
    "pdfplumber","pdfminer","pdfminer.high_level","pdfminer.layout",
    "pywintypes","pythoncom","win32com","win32com.client","win32timezone",
]
hiddenimports += collect_submodules("pdfplumber")
hiddenimports += collect_submodules("pdfminer")
a=Analysis(["TempoFechado_Launcher.py"], pathex=[], binaries=[], datas=datas,
hiddenimports=hiddenimports,
hookspath=[], hooksconfig={}, runtime_hooks=[], excludes=[], win_no_prefer_redirects=False,
win_private_assemblies=False, cipher=block_cipher, noarchive=False)
pyz=PYZ(a.pure, a.zipped_data, cipher=block_cipher)
exe=EXE(pyz, a.scripts, [], exclude_binaries=True, name="TempoFechado", debug=False,
bootloader_ignore_signals=False, strip=False, upx=True, console=False,
icon=str(ROOT/"assets"/"tempo_fechado_app_icon.ico"))
coll=COLLECT(exe, a.binaries, a.zipfiles, a.datas, strip=False, upx=True, upx_exclude=[], name="TempoFechado")
