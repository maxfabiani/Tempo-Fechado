"""
Camada corporativa de usuários - SQLite + senha criptografada.

A senha nunca é gravada em texto puro. O sistema usa os hashes
compatíveis com Werkzeug/Flask, via generate_password_hash e check_password_hash.
"""

from __future__ import annotations

import os
import sqlite3
from datetime import datetime
from pathlib import Path
from typing import Dict, List, Optional, Tuple

from werkzeug.security import check_password_hash, generate_password_hash

PERFIS_VALIDOS = {"admin", "consulta"}

USUARIOS_PADRAO = {} if os.environ.get("TEMPO_FECHADO_SKIP_DEFAULT_USERS") else {
    "admin": {"senha": os.environ.get("TEMPO_FECHADO_ADMIN_PASSWORD") or "admin123", "nome": "Administrador", "perfil": "admin", "trocar_senha": 1},
    "max": {"senha": os.environ.get("TEMPO_FECHADO_MAX_PASSWORD") or "ponto123", "nome": "Max", "perfil": "admin", "trocar_senha": 1},
    "consulta": {"senha": os.environ.get("TEMPO_FECHADO_CONSULTA_PASSWORD") or "consulta123", "nome": "Usuário Consulta", "perfil": "consulta", "trocar_senha": 1},
}


def _conectar(db_path: Path) -> sqlite3.Connection:
    db_path.parent.mkdir(parents=True, exist_ok=True)
    conn = sqlite3.connect(str(db_path))
    conn.row_factory = sqlite3.Row
    return conn


def _agora() -> str:
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")


def _garantir_coluna(conn: sqlite3.Connection, tabela: str, coluna: str, ddl: str) -> None:
    existentes = {r["name"] for r in conn.execute(f"PRAGMA table_info({tabela})").fetchall()}
    if coluna not in existentes:
        conn.execute(f"ALTER TABLE {tabela} ADD COLUMN {coluna} {ddl}")


def _inserir_usuario_inicial(conn: sqlite3.Connection, usuario: str, dados: dict) -> None:
    usuario = (usuario or "").strip().lower()
    if not usuario:
        return
    senha = str(dados.get("senha", "trocar123"))
    perfil = str(dados.get("perfil", "consulta")).lower()
    if perfil not in PERFIS_VALIDOS:
        perfil = "consulta"
    agora = _agora()
    conn.execute(
        """
        INSERT OR IGNORE INTO usuarios
        (usuario, nome, perfil, senha_hash, ativo, trocar_senha, criado_em, atualizado_em, ultimo_login_em)
        VALUES (?, ?, ?, ?, 1, ?, ?, ?, NULL)
        """,
        (
            usuario,
            str(dados.get("nome", usuario)).strip() or usuario,
            perfil,
            generate_password_hash(senha),
            int(dados.get("trocar_senha", 1)),
            agora,
            agora,
        ),
    )


def inicializar_banco_usuarios(db_path: Path, legacy_usuarios: Optional[dict] = None) -> None:
    """Cria o banco/tabela, migra colunas antigas e garante um admin inicial."""
    with _conectar(db_path) as conn:
        conn.execute(
            """
            CREATE TABLE IF NOT EXISTS usuarios (
                usuario TEXT PRIMARY KEY,
                nome TEXT NOT NULL DEFAULT '',
                perfil TEXT NOT NULL DEFAULT 'consulta',
                senha_hash TEXT NOT NULL DEFAULT '',
                ativo INTEGER NOT NULL DEFAULT 1,
                trocar_senha INTEGER NOT NULL DEFAULT 0,
                criado_em TEXT NOT NULL DEFAULT '',
                atualizado_em TEXT NOT NULL DEFAULT '',
                ultimo_login_em TEXT
            )
            """
        )

        # Migração defensiva: se já existir um usuarios.db antigo/incompleto,
        # acrescenta as colunas que a interface de Administração usa.
        _garantir_coluna(conn, "usuarios", "nome", "TEXT NOT NULL DEFAULT ''")
        _garantir_coluna(conn, "usuarios", "perfil", "TEXT NOT NULL DEFAULT 'consulta'")
        _garantir_coluna(conn, "usuarios", "senha_hash", "TEXT NOT NULL DEFAULT ''")
        _garantir_coluna(conn, "usuarios", "ativo", "INTEGER NOT NULL DEFAULT 1")
        _garantir_coluna(conn, "usuarios", "trocar_senha", "INTEGER NOT NULL DEFAULT 0")
        _garantir_coluna(conn, "usuarios", "criado_em", "TEXT NOT NULL DEFAULT ''")
        _garantir_coluna(conn, "usuarios", "atualizado_em", "TEXT NOT NULL DEFAULT ''")
        _garantir_coluna(conn, "usuarios", "ultimo_login_em", "TEXT")

        agora = _agora()
        conn.execute("UPDATE usuarios SET nome = usuario WHERE COALESCE(nome, '') = ''")
        conn.execute("UPDATE usuarios SET perfil = 'consulta' WHERE COALESCE(perfil, '') NOT IN ('admin','consulta')")
        conn.execute("UPDATE usuarios SET ativo = 1 WHERE ativo IS NULL")
        conn.execute("UPDATE usuarios SET trocar_senha = 1 WHERE trocar_senha IS NULL")
        conn.execute("UPDATE usuarios SET criado_em = ? WHERE COALESCE(criado_em, '') = ''", (agora,))
        conn.execute("UPDATE usuarios SET atualizado_em = ? WHERE COALESCE(atualizado_em, '') = ''", (agora,))

        total = conn.execute("SELECT COUNT(*) AS total FROM usuarios").fetchone()["total"]
        if total == 0:
            fonte = legacy_usuarios or USUARIOS_PADRAO
            for usuario, dados in fonte.items():
                _inserir_usuario_inicial(conn, usuario, dados)

        total_admin_ativo = conn.execute(
            "SELECT COUNT(*) AS total FROM usuarios WHERE perfil = 'admin' AND ativo = 1"
        ).fetchone()["total"]
        if total_admin_ativo == 0:
            _inserir_usuario_inicial(conn, "admin", USUARIOS_PADRAO["admin"])

        conn.commit()


def autenticar_usuario(db_path: Path, usuario: str, senha: str) -> Tuple[bool, Optional[Dict], str]:
    usuario = (usuario or "").strip().lower()
    senha = senha or ""
    if not usuario or not senha:
        return False, None, "Usuário e senha são obrigatórios."

    with _conectar(db_path) as conn:
        row = conn.execute("SELECT * FROM usuarios WHERE usuario = ?", (usuario,)).fetchone()
        if not row:
            return False, None, "Usuário ou senha inválidos."
        if int(row["ativo"] or 0) != 1:
            return False, None, "Usuário desativado."
        if not check_password_hash(row["senha_hash"], senha):
            return False, None, "Usuário ou senha inválidos."
        conn.execute("UPDATE usuarios SET ultimo_login_em = ?, atualizado_em = ? WHERE usuario = ?", (_agora(), _agora(), usuario))
        conn.commit()
        return True, _row_para_dict(row), ""


def _row_para_dict(row: sqlite3.Row) -> Dict:
    return {
        "usuario": row["usuario"],
        "nome": row["nome"],
        "perfil": row["perfil"],
        "ativo": bool(row["ativo"]),
        "trocar_senha": bool(row["trocar_senha"]),
        "criado_em": row["criado_em"],
        "atualizado_em": row["atualizado_em"],
        "ultimo_login_em": row["ultimo_login_em"],
    }


def listar_usuarios(db_path: Path) -> List[Dict]:
    with _conectar(db_path) as conn:
        rows = conn.execute(
            """
            SELECT usuario, nome, perfil, ativo, trocar_senha, criado_em, atualizado_em, ultimo_login_em
            FROM usuarios
            ORDER BY perfil ASC, nome COLLATE NOCASE ASC
            """
        ).fetchall()
        return [_row_para_dict(r) for r in rows]


def obter_usuario(db_path: Path, usuario: str) -> Optional[Dict]:
    usuario = (usuario or "").strip().lower()
    if not usuario:
        return None
    with _conectar(db_path) as conn:
        row = conn.execute(
            """
            SELECT usuario, nome, perfil, ativo, trocar_senha, criado_em, atualizado_em, ultimo_login_em
            FROM usuarios
            WHERE usuario = ?
            """,
            (usuario,),
        ).fetchone()
        return _row_para_dict(row) if row else None


def criar_ou_atualizar_usuario(db_path: Path, usuario: str, nome: str, perfil: str, senha: str = "", ativo: bool = True, trocar_senha: bool = True) -> Tuple[bool, str]:
    usuario = (usuario or "").strip().lower()
    nome = (nome or usuario).strip()
    perfil = (perfil or "consulta").strip().lower()

    if not usuario:
        return False, "Informe o usuário."
    if not nome:
        return False, "Informe o nome."
    if perfil not in PERFIS_VALIDOS:
        return False, "Perfil inválido. Use admin ou consulta."

    with _conectar(db_path) as conn:
        row = conn.execute("SELECT usuario FROM usuarios WHERE usuario = ?", (usuario,)).fetchone()
        if row:
            campos = ["nome = ?", "perfil = ?", "ativo = ?", "trocar_senha = ?", "atualizado_em = ?"]
            params = [nome, perfil, int(bool(ativo)), int(bool(trocar_senha)), _agora()]
            if senha:
                campos.append("senha_hash = ?")
                params.append(generate_password_hash(senha))
            params.append(usuario)
            conn.execute(f"UPDATE usuarios SET {', '.join(campos)} WHERE usuario = ?", params)
            conn.commit()
            return True, "Usuário atualizado com sucesso."

        if not senha:
            return False, "Informe uma senha inicial para novo usuário."
        conn.execute(
            """
            INSERT INTO usuarios
            (usuario, nome, perfil, senha_hash, ativo, trocar_senha, criado_em, atualizado_em)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?)
            """,
            (usuario, nome, perfil, generate_password_hash(senha), int(bool(ativo)), int(bool(trocar_senha)), _agora(), _agora()),
        )
        conn.commit()
        return True, "Usuário criado com sucesso."


def alterar_senha_usuario(db_path: Path, usuario: str, nova_senha: str, senha_atual: str = "", exigir_senha_atual: bool = False) -> Tuple[bool, str]:
    usuario = (usuario or "").strip().lower()
    nova_senha = nova_senha or ""
    if len(nova_senha) < 6:
        return False, "A nova senha deve ter pelo menos 6 caracteres."

    with _conectar(db_path) as conn:
        row = conn.execute("SELECT * FROM usuarios WHERE usuario = ?", (usuario,)).fetchone()
        if not row:
            return False, "Usuário não encontrado."
        if exigir_senha_atual and not check_password_hash(row["senha_hash"], senha_atual or ""):
            return False, "Senha atual inválida."
        conn.execute(
            "UPDATE usuarios SET senha_hash = ?, trocar_senha = 0, atualizado_em = ? WHERE usuario = ?",
            (generate_password_hash(nova_senha), _agora(), usuario),
        )
        conn.commit()
        return True, "Senha alterada com sucesso."


def definir_status_usuario(db_path: Path, usuario: str, ativo: bool) -> Tuple[bool, str]:
    usuario = (usuario or "").strip().lower()
    with _conectar(db_path) as conn:
        row = conn.execute("SELECT usuario FROM usuarios WHERE usuario = ?", (usuario,)).fetchone()
        if not row:
            return False, "Usuário não encontrado."
        conn.execute("UPDATE usuarios SET ativo = ?, atualizado_em = ? WHERE usuario = ?", (int(bool(ativo)), _agora(), usuario))
        conn.commit()
        return True, "Status atualizado com sucesso."


def excluir_usuario(db_path: Path, usuario: str) -> Tuple[bool, str]:
    usuario = (usuario or "").strip().lower()
    if not usuario:
        return False, "Informe o usuario."

    with _conectar(db_path) as conn:
        row = conn.execute("SELECT usuario, perfil, ativo FROM usuarios WHERE usuario = ?", (usuario,)).fetchone()
        if not row:
            return False, "Usuario nao encontrado."

        if row["perfil"] == "admin" and int(row["ativo"] or 0) == 1:
            total_admin_ativo = conn.execute(
                "SELECT COUNT(*) AS total FROM usuarios WHERE perfil = 'admin' AND ativo = 1"
            ).fetchone()["total"]
            if total_admin_ativo <= 1:
                return False, "Nao e possivel excluir o ultimo administrador ativo."

        conn.execute("DELETE FROM usuarios WHERE usuario = ?", (usuario,))
        conn.commit()
        return True, "Usuario excluido com sucesso."
