"""
Módulo de conexão e gerenciamento de banco de dados híbrido (Supabase PostgreSQL / SQLite local).
Garante persistência permanente na nuvem no Streamlit Cloud com fallback automático para SQLite.
"""

import copy
import functools
import os
import re
import sqlite3
import time
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


def _clean_pg_url(val: Any) -> Optional[str]:
    """Valida e limpa URL de conexão PostgreSQL. Ignora URLs HTTP/REST."""
    if not val:
        return None
    s = str(val).strip()
    if s.startswith("http://") or s.startswith("https://"):
        return None
    return s


def get_database_url() -> Optional[str]:
    """
    Recupera a URL de conexão do PostgreSQL (Supabase).
    Procura em:
    1. st.secrets ("DATABASE_URL", "SUPABASE_DB_URL", "POSTGRES_URL", "connections.postgresql.url", etc.)
    2. Variáveis de ambiente
    3. Arquivo .streamlit/secrets.toml
    """
    candidate_keys = (
        "DATABASE_URL",
        "SUPABASE_DB_URL",
        "POSTGRES_URL",
        "POSTGRESQL_URL",
        "SUPABASE_POSTGRES_URL",
        "SUPABASE_CONNECTION_STRING",
        "DATABASE_URI",
        "SUPABASE_URL",
    )

    # 1. Streamlit secrets
    try:
        import streamlit as st
        if hasattr(st, "secrets"):
            for k in candidate_keys:
                if k in st.secrets:
                    cleaned = _clean_pg_url(st.secrets[k])
                    if cleaned:
                        return cleaned
            if "connections" in st.secrets:
                for sub in ("postgresql", "postgres", "supabase"):
                    if sub in st.secrets["connections"]:
                        pg_conf = st.secrets["connections"][sub]
                        if isinstance(pg_conf, dict) and "url" in pg_conf:
                            cleaned = _clean_pg_url(pg_conf["url"])
                            if cleaned:
                                return cleaned
    except Exception:
        pass

    # 2. Variáveis de ambiente
    for key in candidate_keys:
        val = os.environ.get(key)
        cleaned = _clean_pg_url(val)
        if cleaned:
            return cleaned

    # 3. Arquivo .streamlit/secrets.toml diretamente
    file_secrets = _get_secrets_from_file()
    for key in candidate_keys:
        if key in file_secrets and file_secrets[key]:
            cleaned = _clean_pg_url(file_secrets[key])
            if cleaned:
                return cleaned
    if "connections" in file_secrets and isinstance(file_secrets["connections"], dict):
        for sub in ("postgresql", "postgres", "supabase"):
            if sub in file_secrets["connections"]:
                pg_c = file_secrets["connections"][sub]
                if isinstance(pg_c, dict) and "url" in pg_c:
                    cleaned = _clean_pg_url(pg_c["url"])
                    if cleaned:
                        return cleaned

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

    def __init__(self, raw_conn, keep_open: bool = True):
        self._conn = raw_conn
        self._keep_open = keep_open

    def cursor(self):
        return PGCursorWrapper(self._conn.cursor())

    def commit(self):
        if not self._conn.closed and not getattr(self._conn, "autocommit", False):
            self._conn.commit()

    def rollback(self):
        if not self._conn.closed and not getattr(self._conn, "autocommit", False):
            self._conn.rollback()

    def close(self):
        if not self._keep_open and not self._conn.closed:
            try:
                self._conn.close()
            except Exception:
                pass

    def __enter__(self):
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        if not self._conn.closed and not getattr(self._conn, "autocommit", False):
            try:
                if exc_type is not None:
                    self.rollback()
                else:
                    self.commit()
            except Exception:
                pass
        self.close()


class QueryCache:
    """Cache em memória para consultas repetitivas de leitura com TTL e invalidação granular."""
    _store: Dict[Any, Tuple[Any, float]] = {}
    DEFAULT_TTL: float = 30.0  # 30 segundos de TTL para consultas de leitura

    @classmethod
    def get(cls, key: Any) -> Optional[Any]:
        if key in cls._store:
            val, expire_at = cls._store[key]
            if time.time() < expire_at:
                return val
            del cls._store[key]
        return None

    @classmethod
    def set(cls, key: Any, val: Any, ttl: Optional[float] = None) -> None:
        expire_at = time.time() + (ttl if ttl is not None else cls.DEFAULT_TTL)
        cls._store[key] = (val, expire_at)

    @classmethod
    def clear(cls) -> None:
        """Limpa todo o cache de consultas."""
        cls._store.clear()

    @classmethod
    def clear_escalas(cls, projeto: Optional[str] = None) -> None:
        """Invalida seletivamente o cache de escalas e períodos do projeto (preserva equipe e configs)."""
        to_del = []
        for k in list(cls._store.keys()):
            if isinstance(k, tuple) and len(k) > 0:
                fn_name = k[0]
                if fn_name in ("get_escalas_periodo", "get_periodos_agrupados", "get_membros_para_escala"):
                    if not projeto:
                        to_del.append(k)
                    else:
                        args_str = " ".join(str(a) for a in k[1])
                        if projeto.upper() in args_str.upper() or "TODOS" in args_str.upper():
                            to_del.append(k)
        for k in to_del:
            cls._store.pop(k, None)

    @classmethod
    def clear_membros(cls) -> None:
        """Invalida o cache de colaboradores e cadastros de equipe."""
        to_del = [
            k for k in list(cls._store.keys())
            if isinstance(k, tuple) and len(k) > 0 and k[0] in ("get_todos_membros", "get_membros_para_escala", "get_regime_colaborador")
        ]
        for k in to_del:
            cls._store.pop(k, None)

    @classmethod
    def clear_sugestoes(cls) -> None:
        """Invalida o cache de sugestões de melhoria."""
        to_del = [
            k for k in list(cls._store.keys())
            if isinstance(k, tuple) and len(k) > 0 and k[0] == "get_sugestoes"
        ]
        for k in to_del:
            cls._store.pop(k, None)

    @classmethod
    def clear_audit_logs(cls) -> None:
        """Invalida o cache do histórico de auditoria."""
        to_del = [
            k for k in list(cls._store.keys())
            if isinstance(k, tuple) and len(k) > 0 and k[0] == "get_audit_logs"
        ]
        for k in to_del:
            cls._store.pop(k, None)

    @classmethod
    def clear_rts(cls) -> None:
        """Invalida o cache de RTs de embarque."""
        to_del = [
            k for k in list(cls._store.keys())
            if isinstance(k, tuple) and len(k) > 0 and k[0] == "get_embarque_rts"
        ]
        for k in to_del:
            cls._store.pop(k, None)

    @classmethod
    def clear_vagas(cls, projeto: Optional[str] = None) -> None:
        """Invalida o cache de metas de vagas mínimas."""
        to_del = []
        for k in list(cls._store.keys()):
            if isinstance(k, tuple) and len(k) > 0 and k[0] == "get_config_vagas":
                if not projeto:
                    to_del.append(k)
                else:
                    args_str = " ".join(str(a) for a in k[1])
                    if projeto.upper() in args_str.upper():
                        to_del.append(k)
        for k in to_del:
            cls._store.pop(k, None)


def cached_query(ttl: float = 30.0):
    """
    Decorador para cache em memória de leituras do banco com TTL e invalidação automática.
    Retorna cópia profunda do resultado para evitar efeitos colaterais de mutação.
    """
    def decorator(fn):
        @functools.wraps(fn)
        def wrapper(*args, **kwargs):
            # Normaliza argumentos ignorando o objeto de classe (cls)
            clean_args = tuple(a for a in args if not isinstance(a, type))
            kw_items = tuple(sorted((k, str(v)) for k, v in kwargs.items()))
            key = (fn.__name__, clean_args, kw_items)

            cached = QueryCache.get(key)
            if cached is not None:
                return copy.deepcopy(cached)

            result = fn(*args, **kwargs)
            QueryCache.set(key, result, ttl=ttl)
            return copy.deepcopy(result)
        return wrapper
    return decorator


class DBManager:
    """
    Gerenciador unificado de banco de dados otimizado para nuvem (Supabase PgBouncer).
    Utiliza conexão persistente direta com autocommit ativado e timeout agressivo,
    eliminando congelamentos de fila (30s/60s) e garantindo resposta sub-segundo.
    """
    _conn = None
    _conn_url: Optional[str] = None
    _initialized: bool = False
    _last_error: Optional[str] = None
    _actual_backend: str = "SQLite (Local)"

    @classmethod
    def is_cloud(cls) -> bool:
        """Verifica se há conexão ativa com Supabase PostgreSQL."""
        return cls._actual_backend == "Supabase (PostgreSQL)"

    @classmethod
    def get_backend_name(cls) -> str:
        """Retorna o nome amigável do backend ativo."""
        return cls._actual_backend

    @classmethod
    def get_diagnostics(cls) -> Dict[str, Any]:
        """Retorna informações diagnósticas do banco de dados para a interface."""
        raw_url = get_database_url()
        return {
            "backend": cls._actual_backend,
            "is_cloud": cls.is_cloud(),
            "url_configured": bool(raw_url),
            "has_psycopg2": HAS_PSYCOPG2,
            "last_error": cls._last_error
        }

    @classmethod
    def clear_cache(cls) -> None:
        """Invalida todo o cache em memória."""
        QueryCache.clear()

    @classmethod
    def get_connection(cls):
        """Retorna conexão ativa persistente com Supabase (autocommit ativo) ou SQLite local."""
        db_url = get_database_url()
        if HAS_PSYCOPG2 and db_url:
            try:
                # Normaliza esquema postgres:// para postgresql:// se necessário
                if db_url.startswith("postgres://"):
                    db_url = "postgresql://" + db_url[len("postgres://"):]

                # Se já temos uma conexão ativa e viva para a mesma URL, reaproveita instantaneamente
                if cls._conn is not None and cls._conn_url == db_url:
                    if not cls._conn.closed:
                        try:
                            cls._conn.poll()
                            cls._actual_backend = "Supabase (PostgreSQL)"
                            cls._last_error = None
                            return PGConnectionWrapper(cls._conn, keep_open=True)
                        except Exception:
                            try:
                                cls._conn.close()
                            except Exception:
                                pass
                            cls._conn = None

                # Abre nova conexão com timeout agressivo de 5s e autocommit=True para PgBouncer
                cls._conn = psycopg2.connect(
                    db_url,
                    connect_timeout=5,
                    keepalives=1,
                    keepalives_idle=30,
                    keepalives_interval=10,
                    keepalives_count=3
                )
                cls._conn.autocommit = True
                cls._conn_url = db_url
                cls._actual_backend = "Supabase (PostgreSQL)"
                cls._last_error = None
                return PGConnectionWrapper(cls._conn, keep_open=True)

            except Exception as e:
                # Se falhar a conexão com a nuvem, loga alerta e faz fallback para SQLite
                cls._last_error = str(e)
                cls._actual_backend = "SQLite (Local)"
                print(f"[DBManager] Alerta: Falha ao conectar ao Supabase ({e}). Utilizando SQLite local.")
                cls._conn = None
        else:
            cls._actual_backend = "SQLite (Local)"

        # Fallback para SQLite local
        conn = sqlite3.connect(DB_PATH)
        conn.row_factory = sqlite3.Row
        return conn

    @classmethod
    def init_db(cls, force: bool = False):
        """
        Cria e atualiza o esquema em PostgreSQL ou SQLite.
        Executado apenas uma vez por inicialização da aplicação para evitar sobrecarga de DDL.
        """
        if cls._initialized and not force:
            return
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
                    regime_folga TEXT DEFAULT '14x21',
                    is_extra INTEGER DEFAULT 0
                )
            """)
            try:
                cursor.execute("ALTER TABLE p80_membros ADD COLUMN is_extra INTEGER DEFAULT 0")
            except Exception:
                pass

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

            # Limpeza de dados de teste (v0.3.2)
            try:
                cursor.execute("DELETE FROM p80_membros WHERE UPPER(nome) LIKE 'TURMA_TESTE%' OR UPPER(nome) LIKE '%TESTE%'")
                cursor.execute("DELETE FROM p80_escalas WHERE UPPER(nome) LIKE 'TURMA_TESTE%' OR UPPER(nome) LIKE '%TESTE%'")
                cursor.execute("DELETE FROM p80_embarque_rt WHERE UPPER(nome) LIKE 'TURMA_TESTE%' OR UPPER(nome) LIKE '%TESTE%'")
                cursor.execute("DELETE FROM p80_audit_log WHERE UPPER(projeto) = 'GERAL' OR UPPER(colaborador) LIKE '%TESTE%' OR UPPER(detalhes) LIKE '%TESTE%' OR UPPER(colaborador) LIKE '%CARLOS SILVA%' OR UPPER(colaborador) LIKE '%MARIANA COSTA%'")
                cursor.execute("DELETE FROM p80_sugestoes WHERE UPPER(autor) LIKE '%CARLOS SILVA%' OR UPPER(autor) LIKE '%MARIANA COSTA%' OR UPPER(titulo) LIKE '%TESTE%'")
            except Exception as e_clean:
                print(f"[DBManager] Aviso na limpeza de dados de teste: {e_clean}")

            conn.commit()

        cls._initialized = True
