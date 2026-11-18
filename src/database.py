"""
Módulo de conexão e gerenciamento de banco de dados híbrido (Supabase PostgreSQL / SQLite local).
Garante persistência permanente na nuvem no Streamlit Cloud com fallback automático para SQLite.
"""

import os
import re
import sqlite3
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

DB_PATH = Path(__file__).parent.parent / "escala_p80.db"

# Tenta carregar biblioteca psycopg2 para conexão com Supabase / PostgreSQL
try:
    import psycopg2
    from psycopg2 import pool
    from psycopg2.extras import RealDictCursor
    HAS_PSYCOPG2 = True
except ImportError:
    HAS_PSYCOPG2 = False


def _get_secrets_from_file() -> Dict[str, Any]:
    """Lê diretamente o arquivo .streamlit/secrets.toml se existir."""
    secrets_file = Path(__file__).parent.parent / ".streamlit" / "secrets.toml"
    if not secrets_file.exists():
        return {}
    try:
        try:
            import tomllib
            with open(secrets_file, "rb") as f:
                return tomllib.load(f)
        except ImportError:
            import toml
            with open(secrets_file, "r", encoding="utf-8") as f:
                return toml.load(f)
    except Exception:
        # Leitura manual simples linha por linha se toml não estiver instalado
        data = {}
        try:
            with open(secrets_file, "r", encoding="utf-8") as f:
                for line in f:
                    line = line.strip()
                    if "=" in line and not line.startswith("#"):
                        k, v = line.split("=", 1)
                        k = k.strip()
                        v = v.strip().strip('"').strip("'")
                        data[k] = v
        except Exception:
            pass
        return data


def get_database_url() -> Optional[str]:
    """
    Recupera a URL de conexão do PostgreSQL (Supabase).
    Procura em:
    1. st.secrets ("DATABASE_URL", "SUPABASE_DB_URL", "connections.postgresql.url")
    2. Variáveis de ambiente
    3. Arquivo .streamlit/secrets.toml
    """
    # 1. Streamlit secrets
    try:
        import streamlit as st
        if hasattr(st, "secrets"):
            if "DATABASE_URL" in st.secrets:
                return str(st.secrets["DATABASE_URL"]).strip()
            if "SUPABASE_DB_URL" in st.secrets:
                return str(st.secrets["SUPABASE_DB_URL"]).strip()
            if "connections" in st.secrets and "postgresql" in st.secrets["connections"]:
                pg_conf = st.secrets["connections"]["postgresql"]
                if isinstance(pg_conf, dict) and "url" in pg_conf:
                    return str(pg_conf["url"]).strip()
    except Exception:
        pass

    # 2. Variáveis de ambiente
    for key in ("DATABASE_URL", "SUPABASE_DB_URL", "POSTGRES_URL"):
        val = os.environ.get(key)
        if val and val.strip():
            return val.strip()

    # 3. Arquivo .streamlit/secrets.toml diretamente
    file_secrets = _get_secrets_from_file()
    for key in ("DATABASE_URL", "SUPABASE_DB_URL", "POSTGRES_URL"):
        if key in file_secrets and file_secrets[key]:
            return str(file_secrets[key]).strip()
    if "connections" in file_secrets and isinstance(file_secrets["connections"], dict):
        pg_c = file_secrets["connections"].get("postgresql", {})
        if isinstance(pg_c, dict) and "url" in pg_c:
            return str(pg_c["url"]).strip()

    return None


class RowWrapper:
    """Wrapper para compatibilizar resultados de PostgreSQL com sqlite3.Row."""

    def __init__(self, cols: List[str], values: Tuple[Any, ...]):
        self._cols = [c.lower() for c in cols]
        self._col_to_idx = {c: i for i, c in enumerate(self._cols)}
        self._values = tuple(values)

    def __getitem__(self, key: Any) -> Any:
        if isinstance(key, int):
            return self._values[key]
        idx = self._col_to_idx.get(str(key).lower())
        if idx is not None:
            return self._values[idx]
        raise KeyError(key)

    def get(self, key: str, default: Any = None) -> Any:
        idx = self._col_to_idx.get(str(key).lower())
        return self._values[idx] if idx is not None else default

    def keys(self) -> List[str]:
        return list(self._cols)

    def __iter__(self):
        return iter(self._values)

    def __len__(self) -> int:
        return len(self._values)

    def __repr__(self) -> str:
        d = {c: v for c, v in zip(self._cols, self._values)}
        return f"<Row {d}>"


def _convert_sql_placeholders(sql: str) -> str:
    """Substitui marcadores '?' por '%s' fora de strings literais para PostgreSQL."""
    parts = sql.split("'")
    for i in range(0, len(parts), 2):
        parts[i] = parts[i].replace("?", "%s")
    return "'".join(parts)


class PGCursorWrapper:
    """Cursor wrapper que adapta queries '?' para '%s' e constrói RowWrappers."""

    def __init__(self, raw_cursor):
        self._cursor = raw_cursor

    def execute(self, sql: str, params: Optional[Any] = None):
        pg_sql = _convert_sql_placeholders(sql)
        if params is not None:
            if isinstance(params, (list, tuple)):
                # Converte parâmetros para tupla compatível com psycopg2
                params = tuple(params)
            return self._cursor.execute(pg_sql, params)
        return self._cursor.execute(pg_sql)

    def executemany(self, sql: str, param_seq: List[Any]):
        pg_sql = _convert_sql_placeholders(sql)
        return self._cursor.executemany(pg_sql, param_seq)

    def fetchone(self) -> Optional[RowWrapper]:
        row = self._cursor.fetchone()
        if row is None:
            return None
        cols = [desc[0] for desc in self._cursor.description]
        return RowWrapper(cols, row)

    def fetchall(self) -> List[RowWrapper]:
        rows = self._cursor.fetchall()
        if not rows:
            return []
        cols = [desc[0] for desc in self._cursor.description]
        return [RowWrapper(cols, r) for r in rows]

    def fetchmany(self, size: int) -> List[RowWrapper]:
        rows = self._cursor.fetchmany(size)
        if not rows:
            return []
        cols = [desc[0] for desc in self._cursor.description]
        return [RowWrapper(cols, r) for r in rows]

    @property
    def rowcount(self) -> int:
        return self._cursor.rowcount

    def close(self):
        self._cursor.close()

    def __enter__(self):
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        self.close()


class PGConnectionWrapper:
    """Wrapper para conexão PostgreSQL do psycopg2 compatível com interface de conexão do SQLite."""

    def __init__(self, raw_conn):
        self._conn = raw_conn

    def cursor(self):
        return PGCursorWrapper(self._conn.cursor())

    def commit(self):
        self._conn.commit()

    def rollback(self):
        self._conn.rollback()

    def close(self):
        self._conn.close()

    def __enter__(self):
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        if exc_type is not None:
            self._conn.rollback()
        else:
            self._conn.commit()
        self._conn.close()


class DBManager:
    """
    Gerenciador unificado de banco de dados.
    Conecta automaticamente ao Supabase (PostgreSQL) se a URL estiver configurada;
    caso contrário, utiliza o banco SQLite local (escala_p80.db).
    """

    @classmethod
    def is_cloud(cls) -> bool:
        """Verifica se há configuração ativa para Supabase PostgreSQL."""
        return HAS_PSYCOPG2 and bool(get_database_url())

    @classmethod
    def get_backend_name(cls) -> str:
        """Retorna o nome amigável do backend ativo."""
        return "Supabase (PostgreSQL)" if cls.is_cloud() else "SQLite (Local)"

    @classmethod
    def get_connection(cls):
        """Retorna conexão ativa com o banco configurado (Supabase ou SQLite local)."""
        db_url = get_database_url()
        if HAS_PSYCOPG2 and db_url:
            try:
                # Normaliza esquema postgres:// para postgresql:// se necessário
                if db_url.startswith("postgres://"):
                    db_url = "postgresql://" + db_url[len("postgres://"):]
                raw_conn = psycopg2.connect(db_url)
                return PGConnectionWrapper(raw_conn)
            except Exception as e:
                # Se falhar a conexão com a nuvem, loga alerta e faz fallback para SQLite
                print(f"[DBManager] Alerta: Falha ao conectar ao Supabase ({e}). Utilizando SQLite local.")

        # Fallback para SQLite local
        conn = sqlite3.connect(DB_PATH)
        conn.row_factory = sqlite3.Row
        return conn

    @classmethod
    def init_db(cls):
        """Cria e atualiza o esquema em PostgreSQL ou SQLite."""
        is_pg = cls.is_cloud()

        with cls.get_connection() as conn:
            cursor = conn.cursor()

            id_col = "id SERIAL PRIMARY KEY" if is_pg else "id INTEGER PRIMARY KEY AUTOINCREMENT"

            # 1. Tabela de Membros da Equipe
            cursor.execute(f"""
                CREATE TABLE IF NOT EXISTS p80_membros (
                    {id_col},
                    nome TEXT UNIQUE NOT NULL,
                    disciplina TEXT NOT NULL,
                    funcao TEXT,
                    projeto TEXT DEFAULT 'P-80',
                    ativo INTEGER DEFAULT 1,
                    ordem INTEGER DEFAULT 0,
                    regime_folga TEXT DEFAULT '14x21'
                )
            """)

            # 2. Tabela de Escalas de Plataforma
            cursor.execute(f"""
                CREATE TABLE IF NOT EXISTS p80_escalas (
                    {id_col},
                    nome TEXT NOT NULL,
                    disciplina TEXT NOT NULL,
                    projeto TEXT NOT NULL,
                    data TEXT NOT NULL,
                    status TEXT NOT NULL,
                    observacao TEXT,
                    updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                    UNIQUE(nome, projeto, data)
                )
            """)
            cursor.execute("CREATE INDEX IF NOT EXISTS idx_escalas_proj_data ON p80_escalas(projeto, data)")
            cursor.execute("CREATE INDEX IF NOT EXISTS idx_escalas_nome ON p80_escalas(nome)")

            # 3. Tabela de Sugestões
            cursor.execute(f"""
                CREATE TABLE IF NOT EXISTS p80_sugestoes (
                    {id_col},
                    autor TEXT NOT NULL,
                    is_anonimo INTEGER DEFAULT 0,
                    titulo TEXT NOT NULL,
                    descricao TEXT NOT NULL,
                    prioridade TEXT DEFAULT 'Média',
                    status TEXT DEFAULT 'Pendente',
                    created_at TEXT NOT NULL,
                    updated_at TEXT NOT NULL,
                    implementada_at TEXT
                )
            """)
            cursor.execute("CREATE INDEX IF NOT EXISTS idx_sugestoes_status ON p80_sugestoes(status)")
            cursor.execute("CREATE INDEX IF NOT EXISTS idx_sugestoes_prio ON p80_sugestoes(prioridade)")

            # 4. Tabela de Log de Auditoria
            cursor.execute(f"""
                CREATE TABLE IF NOT EXISTS p80_audit_log (
                    {id_col},
                    data_hora TEXT NOT NULL,
                    usuario TEXT NOT NULL,
                    acao TEXT NOT NULL,
                    projeto TEXT DEFAULT 'GERAL',
                    colaborador TEXT DEFAULT '—',
                    detalhes TEXT NOT NULL
                )
            """)
            if is_pg:
                cursor.execute("CREATE INDEX IF NOT EXISTS idx_audit_log_data ON p80_audit_log(data_hora DESC)")
            else:
                cursor.execute("CREATE INDEX IF NOT EXISTS idx_audit_log_data ON p80_audit_log(data_hora)")
            cursor.execute("CREATE INDEX IF NOT EXISTS idx_audit_log_proj ON p80_audit_log(projeto)")

            # 5. Tabela de Vagas Mínimas
            cursor.execute("""
                CREATE TABLE IF NOT EXISTS p80_config_vagas (
                    projeto TEXT PRIMARY KEY,
                    min_mecanica INTEGER DEFAULT 0,
                    min_instrumentacao INTEGER DEFAULT 0,
                    min_eletrica INTEGER DEFAULT 0,
                    min_total INTEGER DEFAULT 0,
                    updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
                )
            """)

            # 6. Tabela de RTs de Embarques
            cursor.execute(f"""
                CREATE TABLE IF NOT EXISTS p80_embarque_rt (
                    {id_col},
                    projeto TEXT NOT NULL,
                    nome TEXT NOT NULL,
                    data_inicio TEXT NOT NULL,
                    data_fim TEXT NOT NULL,
                    numero_rt TEXT DEFAULT '',
                    updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                    UNIQUE(projeto, nome, data_inicio, data_fim)
                )
            """)

            # Normalizações e atualizações de integridade
            cursor.execute("UPDATE p80_membros SET disciplina = 'Elétrica' WHERE disciplina LIKE '%TRIC%' OR disciplina LIKE '%ELET%'")
            cursor.execute("UPDATE p80_escalas SET disciplina = 'Elétrica' WHERE disciplina LIKE '%TRIC%' OR disciplina LIKE '%ELET%'")
            cursor.execute("UPDATE p80_membros SET disciplina = 'Mecânica' WHERE disciplina LIKE '%MEC%'")
            cursor.execute("UPDATE p80_escalas SET disciplina = 'Mecânica' WHERE disciplina LIKE '%MEC%'")
            cursor.execute("UPDATE p80_membros SET disciplina = 'Instrumentação' WHERE disciplina LIKE '%INST%'")
            cursor.execute("UPDATE p80_escalas SET disciplina = 'Instrumentação' WHERE disciplina LIKE '%INST%'")

            cursor.execute("UPDATE p80_escalas SET status = 'EMB' WHERE status = 'E'")
            cursor.execute("UPDATE p80_membros SET ativo = 0 WHERE UPPER(funcao) LIKE '%SUPERVIS%' OR UPPER(funcao) LIKE '%COORD%'")
            conn.commit()
