import os, sys, time, socket, threading, traceback, webbrowser, shutil, json, urllib.request, subprocess, runpy
from pathlib import Path

HOST="127.0.0.1"
PORT=5050
URL=f"http://{HOST}:{PORT}"
APP_VERSION="v8.21.38"
APP_RELEASE_NAME="Upload Manual PDFs"
APP_FULL_NAME="Tempo Fechado v8.21.38 - Upload Manual PDFs"
PASTAS=["entrada","saida","processados","logs","backup","erro","config","controle","backup_atualizacao"]

def base_dir(): return Path(sys.executable).resolve().parent if getattr(sys,"frozen",False) else Path(__file__).resolve().parent

def log_file_path():
    try:
        base = Path.home() / "ponto_pdfs" / "logs"
        base.mkdir(parents=True, exist_ok=True)
        return base / "tempo_fechado_launcher.log"
    except Exception:
        return base_dir() / "tempo_fechado_launcher.log"

def log(msg):
    linha = f"[TempoFechado] {time.strftime('%Y-%m-%d %H:%M:%S')} - {msg}"
    try:
        with open(log_file_path(), "a", encoding="utf-8") as f:
            f.write(linha + "\n")
    except Exception:
        pass
    # Em modo console ainda permite acompanhar; em modo --noconsole simplesmente nao aparece janela.
    try:
        if sys.stdout:
            print(linha, flush=True)
    except Exception:
        pass

def log_exception(prefix):
    try:
        tb = traceback.format_exc()
        log(prefix)
        with open(log_file_path(), "a", encoding="utf-8") as f:
            f.write(tb + "\n")
    except Exception:
        pass

def avisar_erro_sem_console(titulo, mensagem):
    """Mostra erro em janela quando o EXE roda sem console."""
    try:
        if os.name == "nt" and getattr(sys, "frozen", False):
            import ctypes
            ctypes.windll.user32.MessageBoxW(None, mensagem, titulo, 0x10)
    except Exception:
        pass

def garantir_ponto_pdfs():
    base = Path.home() / "ponto_pdfs"
    base.mkdir(parents=True, exist_ok=True)
    for p in PASTAS:
        (base / p).mkdir(parents=True, exist_ok=True)
    log(f"Estrutura ponto_pdfs validada/criada: {base}")
    return base

def _copiar_se_existir(origem: Path, destino: Path, manifesto: list):
    try:
        if not origem.exists():
            return
        destino.parent.mkdir(parents=True, exist_ok=True)
        if origem.is_dir():
            shutil.copytree(origem, destino, dirs_exist_ok=True)
            manifesto.append({"tipo": "pasta", "origem": str(origem), "destino": str(destino)})
        else:
            shutil.copy2(origem, destino)
            manifesto.append({"tipo": "arquivo", "origem": str(origem), "destino": str(destino)})
    except Exception as e:
        manifesto.append({"tipo": "erro", "origem": str(origem), "erro": str(e)})

def criar_backup_atualizacao_segura_v8201():
    """Cria uma copia defensiva dos dados locais antes de iniciar o app.

    A instalacao por usuario normalmente preserva arquivos criados localmente,
    mas esta rotina deixa um paraquedas extra em ~/ponto_pdfs/backup_atualizacao.
    """
    try:
        base = garantir_ponto_pdfs()
        app_base = base_dir()
        backup_raiz = base / "backup_atualizacao"
        marcador = backup_raiz / "ultimo_backup_v8201.txt"

        # Evita criar varios backups no mesmo dia em cada abertura do aplicativo.
        hoje = time.strftime("%Y-%m-%d")
        if marcador.exists() and marcador.read_text(encoding="utf-8", errors="ignore").strip() == hoje:
            log("Backup de atualizacao segura ja existe para hoje.")
            return

        destino = backup_raiz / f"backup_{time.strftime('%Y%m%d_%H%M%S')}"
        manifesto = []

        # Dados gerados pelo app na pasta instalada.
        for nome in ["data", "usuarios.db", "tempo_fechado.db", "config.json", "config_robo.json", "alertas_notificacao_config.json"]:
            _copiar_se_existir(app_base / nome, destino / "app" / nome, manifesto)

        # Configuracoes e logs persistentes do perfil do usuario.
        for nome in ["config", "logs", "controle"]:
            _copiar_se_existir(base / nome, destino / "ponto_pdfs" / nome, manifesto)

        if manifesto:
            destino.mkdir(parents=True, exist_ok=True)
            (destino / "MANIFESTO_ATUALIZACAO_SEGURA.json").write_text(
                json.dumps({
                    "versao": APP_FULL_NAME,
                    "gerado_em": time.strftime("%Y-%m-%d %H:%M:%S"),
                    "itens": manifesto,
                }, ensure_ascii=False, indent=2),
                encoding="utf-8"
            )
            marcador.write_text(hoje, encoding="utf-8")
            log(f"Backup de atualizacao segura criado: {destino}")
        else:
            log("Nenhum dado local encontrado para backup de atualizacao segura.")
    except Exception as e:
        log(f"Aviso: falha ao criar backup de atualizacao segura: {e}")

def porta_aberta():
    try:
        with socket.create_connection((HOST, PORT), timeout=0.5):
            return True
    except OSError:
        return False



def pids_escutando_porta_v8217():
    """Retorna PIDs que estao em LISTENING na porta do Tempo Fechado.

    Esta rotina existe porque, no modo --noconsole, uma instancia antiga pode
    continuar viva em segundo plano mesmo depois do navegador ser fechado.
    """
    pids = set()
    if os.name != "nt":
        return pids
    try:
        cmd = f'netstat -ano | findstr :{PORT}'
        proc = subprocess.run(cmd, shell=True, capture_output=True, text=True, timeout=5)
        for linha in (proc.stdout or "").splitlines():
            partes = linha.split()
            if len(partes) < 5:
                continue
            protocolo = partes[0].upper()
            endereco_local = partes[1]
            estado = partes[3].upper() if protocolo == "TCP" else ""
            pid = partes[-1]
            if protocolo == "TCP" and endereco_local.endswith(f":{PORT}") and estado == "LISTENING" and pid.isdigit():
                if int(pid) != os.getpid():
                    pids.add(int(pid))
    except Exception as e:
        log(f"Aviso: nao foi possivel consultar PIDs da porta {PORT}: {e}")
    return pids


def encerrar_instancias_antigas_v8217(pids):
    """Finaliza processos antigos segurando a porta 5050 no Windows."""
    encerrados = []
    if os.name != "nt":
        return encerrados
    for pid in sorted(set(pids)):
        try:
            log(f"Encerrando instancia antiga na porta {PORT}. PID={pid}")
            subprocess.run(["taskkill", "/PID", str(pid), "/F"], capture_output=True, text=True, timeout=10)
            encerrados.append(pid)
        except Exception as e:
            log(f"Falha ao encerrar PID {pid}: {e}")
    # Aguarda a porta liberar. Estados TIME_WAIT podem continuar, mas LISTENING deve sumir.
    t0 = time.time()
    while time.time() - t0 < 10:
        if not pids_escutando_porta_v8217():
            return encerrados
        time.sleep(0.5)
    return encerrados


def aguardar(timeout=120):
    log(f"Aguardando Flask em {URL} por ate {timeout}s...")
    t0=time.time()
    aviso=-1
    while time.time()-t0 < timeout:
        if porta_aberta():
            log("Flask respondeu.")
            return True
        e=int(time.time()-t0)
        if e and e%10==0 and e!=aviso:
            aviso=e
            log("Ainda aguardando servidor...")
        time.sleep(1)
    return False

def consultar_versao_instancia_ativa_v8215():
    """Consulta uma instancia ja aberta na porta 5050.

    Quando uma instalacao antiga fica rodando, o novo EXE apenas abre o navegador
    nessa instancia antiga. Esta checagem deixa isso visivel no log e, no modo EXE,
    em uma mensagem simples para o operador.
    """
    try:
        with urllib.request.urlopen(f"{URL}/api/versao-publica", timeout=2) as resp:
            texto = resp.read().decode("utf-8", errors="ignore")
            dados = json.loads(texto)
            return dados
    except Exception as e:
        return {"ok": False, "erro": str(e)}


def avisar_instancia_antiga_v8215(info):
    try:
        versao = info.get("versao") if isinstance(info, dict) else None
        if versao == APP_VERSION:
            return
        detalhe = ""
        if isinstance(info, dict):
            detalhe = info.get("versao") or info.get("erro") or str(info)
        mensagem = (
            "Já existe uma instância do Tempo Fechado aberta na porta 5050.\n\n"
            f"Launcher atual: {APP_FULL_NAME}\n"
            f"Instância encontrada: {detalhe or 'não identificada'}\n\n"
            "Se a tela continuar mostrando versão antiga, feche o Tempo Fechado, finalize processos "
            "TempoFechado.exe/python.exe/pythonw.exe no Gerenciador de Tarefas e abra novamente."
        )
        log("Possivel instancia antiga detectada: " + (detalhe or "sem versao"))
        avisar_erro_sem_console("Tempo Fechado - versão ativa", mensagem)
    except Exception:
        pass


def iniciar_servidor():
    try:
        os.chdir(base_dir())
        garantir_ponto_pdfs()
        log(f"Pasta base do app: {base_dir()}")
        log("Importando robo_ponto_web...")
        import robo_ponto_web
        app = getattr(robo_ponto_web, "app", None)
        if app is None:
            raise RuntimeError("Nao encontrei a variavel Flask app em robo_ponto_web.py")
        log("Iniciando Flask...")
        app.run(debug=False, host="0.0.0.0", port=PORT, threaded=True, use_reloader=False)
    except Exception:
        log_exception("ERRO ao iniciar Flask:")
        avisar_erro_sem_console("Tempo Fechado", f"Erro ao iniciar o servidor. Consulte o log:\n{log_file_path()}")
        raise

def executar_script_embarcado_v82124():
    script = os.environ.get("TEMPO_FECHADO_RUN_SCRIPT", "").strip().strip('"').strip("'")
    if not script:
        return False
    try:
        script_path = Path(script).resolve()
        cwd = os.environ.get("TEMPO_FECHADO_RUN_SCRIPT_CWD", "").strip()
        if cwd:
            os.chdir(cwd)
        else:
            os.chdir(str(script_path.parent))
        sys.argv = [str(script_path)]
        script_dir = str(script_path.parent)
        if script_dir not in sys.path:
            sys.path.insert(0, script_dir)
        log(f"Executando script embarcado pelo TempoFechado.exe: {script_path}")
        runpy.run_path(str(script_path), run_name="__main__")
        return True
    except SystemExit as e:
        raise e
    except Exception:
        log_exception("ERRO ao executar script embarcado:")
        traceback.print_exc(file=sys.stderr)
        raise SystemExit(1)

def main():
    try:
        if executar_script_embarcado_v82124():
            return
        log("Inicializando Tempo Fechado v8.21.38 - Upload Manual PDFs...")
        garantir_ponto_pdfs()
        criar_backup_atualizacao_segura_v8201()
        if porta_aberta():
            info = consultar_versao_instancia_ativa_v8215()
            versao_ativa = info.get("versao") if isinstance(info, dict) else None
            if versao_ativa == APP_VERSION:
                log(f"Instancia atual ja ativa em {URL}: {info.get('titulo') or versao_ativa}")
                webbrowser.open(URL)
                return

            if isinstance(info, dict) and info.get("versao"):
                log(f"Instancia diferente encontrada em {URL}: {info.get('titulo') or info.get('versao')}")
            else:
                log(f"Instancia antiga/nao identificada em {URL}: {info}")

            pids = pids_escutando_porta_v8217()
            if pids:
                encerrados = encerrar_instancias_antigas_v8217(pids)
                log(f"Instancias antigas encerradas automaticamente: {encerrados}")
                if porta_aberta():
                    avisar_instancia_antiga_v8215(info if isinstance(info, dict) else {})
                    webbrowser.open(URL)
                    return
            else:
                avisar_instancia_antiga_v8215(info if isinstance(info, dict) else {})
                webbrowser.open(URL)
                return
        th=threading.Thread(target=iniciar_servidor, daemon=True)
        th.start()
        if aguardar():
            webbrowser.open(URL)
            log("Tempo Fechado iniciado em modo sem console. Logs em ~/ponto_pdfs/logs/tempo_fechado_launcher.log.")
            while True: time.sleep(1)
        raise RuntimeError(f"Tempo Fechado nao respondeu em {URL}")
    except Exception:
        log_exception("ERRO GERAL:")
        avisar_erro_sem_console("Tempo Fechado", f"Erro ao abrir o Tempo Fechado. Consulte o log:\n{log_file_path()}")
        raise

if __name__=="__main__":
    main()

