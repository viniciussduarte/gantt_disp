"""
Script utilitário para migrar todos os dados do banco local SQLite (escala_p80.db)
para o Supabase (PostgreSQL).

Uso:
    python migrate_sqlite_to_supabase.py [DATABASE_URL]

Se DATABASE_URL não for informado via linha de comando, procurará em .streamlit/secrets.toml
ou na variável de ambiente DATABASE_URL.
"""

import os
import sqlite3
import sys
from pathlib import Path

# Adiciona o diretório raiz ao path
ROOT_DIR = Path(__file__).parent
sys.path.insert(0, str(ROOT_DIR))

from src.database import DB_PATH, DBManager, get_database_url


def run_migration(pg_url: str = None):
    url = pg_url or get_database_url()
    if not url:
        print("\n❌ ERRO: Nenhuma URL de conexão do Supabase foi fornecida!")
        print("Forneça a URL como argumento ou configure em .streamlit/secrets.toml:")
        print("  DATABASE_URL = \"postgresql://postgres.[ref]:[senha]@aws-0-[regiao].pooler.supabase.com:6543/postgres\"\n")
        return False

    if not Path(DB_PATH).exists():
        print(f"\n❌ ERRO: Arquivo SQLite local não encontrado em: {DB_PATH}")
        return False

    try:
        import psycopg2
    except ImportError:
        print("\n❌ ERRO: psycopg2-binary não instalado. Execute: pip install psycopg2-binary")
        return False

    print("=" * 70)
    print("🚀 INICIANDO MIGRAÇÃO: SQLite local -> Supabase (PostgreSQL)")
    print("=" * 70)
    print(f"📁 Origem (SQLite): {DB_PATH}")
    # Oculta a senha na exibição da URL
    masked_url = url
    if "@" in url:
        prefix, host_part = url.split("@", 1)
        if ":" in prefix:
            proto, rest = prefix.split("://", 1) if "://" in prefix else ("postgres", prefix)
            user = rest.split(":")[0] if ":" in rest else rest
            masked_url = f"{proto}://{user}:*****@{host_part}"
    print(f"☁️ Destino (Supabase): {masked_url}")

    # Normaliza esquema se necessário
    if url.startswith("postgres://"):
        url = "postgresql://" + url[len("postgres://"):]

    try:
        pg_conn = psycopg2.connect(url)
        pg_conn.autocommit = False
    except Exception as e:
        print(f"\n❌ ERRO ao conectar ao Supabase: {e}")
        return False

    sqlite_conn = sqlite3.connect(DB_PATH)
    sqlite_conn.row_factory = sqlite3.Row

    try:
        # 1. Inicializa o esquema de tabelas no Supabase
        print("\n1️⃣ Criando tabelas e índices no Supabase...")
        with pg_conn.cursor() as cur:
            cur.execute("""
                CREATE TABLE IF NOT EXISTS p80_membros (
                    id SERIAL PRIMARY KEY,
                    nome TEXT UNIQUE NOT NULL,
                    disciplina TEXT NOT NULL,
                    funcao TEXT,
                    projeto TEXT DEFAULT 'P-80',
                    ativo INTEGER DEFAULT 1,
                    ordem INTEGER DEFAULT 0,
                    regime_folga TEXT DEFAULT '14x21',
                    is_extra INTEGER DEFAULT 0
                );
                CREATE TABLE IF NOT EXISTS p80_escalas (
                    id SERIAL PRIMARY KEY,
                    nome TEXT NOT NULL,
                    disciplina TEXT NOT NULL,
                    projeto TEXT NOT NULL,
                    data TEXT NOT NULL,
                    status TEXT NOT NULL,
                    observacao TEXT,
                    updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                    UNIQUE(nome, projeto, data)
                );
                CREATE INDEX IF NOT EXISTS idx_escalas_proj_data ON p80_escalas(projeto, data);
                CREATE INDEX IF NOT EXISTS idx_escalas_nome ON p80_escalas(nome);

                CREATE TABLE IF NOT EXISTS p80_sugestoes (
                    id SERIAL PRIMARY KEY,
                    autor TEXT NOT NULL,
                    is_anonimo INTEGER DEFAULT 0,
                    titulo TEXT NOT NULL,
                    descricao TEXT NOT NULL,
                    prioridade TEXT DEFAULT 'Média',
                    status TEXT DEFAULT 'Pendente',
                    created_at TEXT NOT NULL,
                    updated_at TEXT NOT NULL,
                    implementada_at TEXT
                );
                CREATE INDEX IF NOT EXISTS idx_sugestoes_status ON p80_sugestoes(status);
                CREATE INDEX IF NOT EXISTS idx_sugestoes_prio ON p80_sugestoes(prioridade);

                CREATE TABLE IF NOT EXISTS p80_audit_log (
                    id SERIAL PRIMARY KEY,
                    data_hora TEXT NOT NULL,
                    usuario TEXT NOT NULL,
                    acao TEXT NOT NULL,
                    projeto TEXT DEFAULT 'GERAL',
                    colaborador TEXT DEFAULT '—',
                    detalhes TEXT NOT NULL
                );
                CREATE INDEX IF NOT EXISTS idx_audit_log_data ON p80_audit_log(data_hora DESC);
                CREATE INDEX IF NOT EXISTS idx_audit_log_proj ON p80_audit_log(projeto);

                CREATE TABLE IF NOT EXISTS p80_config_vagas (
                    projeto TEXT PRIMARY KEY,
                    min_mecanica INTEGER DEFAULT 0,
                    min_instrumentacao INTEGER DEFAULT 0,
                    min_eletrica INTEGER DEFAULT 0,
                    min_total INTEGER DEFAULT 0,
                    updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
                );

                CREATE TABLE IF NOT EXISTS p80_embarque_rt (
                    id SERIAL PRIMARY KEY,
                    projeto TEXT NOT NULL,
                    nome TEXT NOT NULL,
                    data_inicio TEXT NOT NULL,
                    data_fim TEXT NOT NULL,
                    numero_rt TEXT DEFAULT '',
                    updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                    UNIQUE(projeto, nome, data_inicio, data_fim)
                );
            """)
        pg_conn.commit()
        print("   ✅ Esquema verificado com sucesso.")

        # 2. Migra p80_membros
        print("\n2️⃣ Migrando 'p80_membros'...")
        cur_sq = sqlite_conn.cursor()
        cur_sq.execute("SELECT nome, disciplina, funcao, projeto, ativo, ordem, regime_folga, COALESCE(is_extra, 0) as is_extra FROM p80_membros")
        membros_rows = cur_sq.fetchall()
        with pg_conn.cursor() as cur:
            for r in membros_rows:
                cur.execute("""
                    INSERT INTO p80_membros (nome, disciplina, funcao, projeto, ativo, ordem, regime_folga, is_extra)
                    VALUES (%s, %s, %s, %s, %s, %s, %s, %s)
                    ON CONFLICT (nome) DO UPDATE SET
                        disciplina = EXCLUDED.disciplina,
                        funcao = EXCLUDED.funcao,
                        projeto = EXCLUDED.projeto,
                        ativo = EXCLUDED.ativo,
                        ordem = EXCLUDED.ordem,
                        regime_folga = EXCLUDED.regime_folga,
                        is_extra = EXCLUDED.is_extra
                """, (r["nome"], r["disciplina"], r["funcao"], r["projeto"], r["ativo"], r["ordem"], r["regime_folga"] or "14x21", r["is_extra"]))
        pg_conn.commit()
        print(f"   ✅ {len(membros_rows)} colaboradores migrados.")

        # 3. Migra p80_escalas
        print("\n3️⃣ Migrando 'p80_escalas'...")
        cur_sq.execute("SELECT nome, disciplina, projeto, data, status, observacao, updated_at FROM p80_escalas")
        escalas_rows = cur_sq.fetchall()
        with pg_conn.cursor() as cur:
            for r in escalas_rows:
                cur.execute("""
                    INSERT INTO p80_escalas (nome, disciplina, projeto, data, status, observacao, updated_at)
                    VALUES (%s, %s, %s, %s, %s, %s, %s)
                    ON CONFLICT (nome, projeto, data) DO UPDATE SET
                        disciplina = EXCLUDED.disciplina,
                        status = EXCLUDED.status,
                        observacao = EXCLUDED.observacao,
                        updated_at = EXCLUDED.updated_at
                """, (r["nome"], r["disciplina"], r["projeto"], r["data"], r["status"], r["observacao"], r["updated_at"]))
        pg_conn.commit()
        print(f"   ✅ {len(escalas_rows)} escalas diárias migradas.")

        # 4. Migra p80_sugestoes
        print("\n4️⃣ Migrando 'p80_sugestoes'...")
        cur_sq.execute("SELECT id, autor, is_anonimo, titulo, descricao, prioridade, status, created_at, updated_at, implementada_at FROM p80_sugestoes")
        sug_rows = cur_sq.fetchall()
        with pg_conn.cursor() as cur:
            for r in sug_rows:
                cur.execute("""
                    INSERT INTO p80_sugestoes (id, autor, is_anonimo, titulo, descricao, prioridade, status, created_at, updated_at, implementada_at)
                    VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s)
                    ON CONFLICT (id) DO UPDATE SET
                        autor = EXCLUDED.autor,
                        is_anonimo = EXCLUDED.is_anonimo,
                        titulo = EXCLUDED.titulo,
                        descricao = EXCLUDED.descricao,
                        prioridade = EXCLUDED.prioridade,
                        status = EXCLUDED.status,
                        created_at = EXCLUDED.created_at,
                        updated_at = EXCLUDED.updated_at,
                        implementada_at = EXCLUDED.implementada_at
                """, (r["id"], r["autor"], r["is_anonimo"], r["titulo"], r["descricao"], r["prioridade"], r["status"], r["created_at"], r["updated_at"], r["implementada_at"]))
            cur.execute("SELECT setval(pg_get_serial_sequence('p80_sugestoes', 'id'), coalesce(max(id), 1)) FROM p80_sugestoes")
        pg_conn.commit()
        print(f"   ✅ {len(sug_rows)} sugestões migradas.")

        # 5. Migra p80_audit_log
        print("\n5️⃣ Migrando 'p80_audit_log'...")
        cur_sq.execute("SELECT id, data_hora, usuario, acao, projeto, colaborador, detalhes FROM p80_audit_log")
        log_rows = cur_sq.fetchall()
        with pg_conn.cursor() as cur:
            for r in log_rows:
                cur.execute("""
                    INSERT INTO p80_audit_log (id, data_hora, usuario, acao, projeto, colaborador, detalhes)
                    VALUES (%s, %s, %s, %s, %s, %s, %s)
                    ON CONFLICT (id) DO NOTHING
                """, (r["id"], r["data_hora"], r["usuario"], r["acao"], r["projeto"], r["colaborador"], r["detalhes"]))
            cur.execute("SELECT setval(pg_get_serial_sequence('p80_audit_log', 'id'), coalesce(max(id), 1)) FROM p80_audit_log")
        pg_conn.commit()
        print(f"   ✅ {len(log_rows)} logs de auditoria migrados.")

        # 6. Migra p80_config_vagas
        print("\n6️⃣ Migrando 'p80_config_vagas'...")
        cur_sq.execute("SELECT projeto, min_mecanica, min_instrumentacao, min_eletrica, min_total, updated_at FROM p80_config_vagas")
        vagas_rows = cur_sq.fetchall()
        with pg_conn.cursor() as cur:
            for r in vagas_rows:
                cur.execute("""
                    INSERT INTO p80_config_vagas (projeto, min_mecanica, min_instrumentacao, min_eletrica, min_total, updated_at)
                    VALUES (%s, %s, %s, %s, %s, %s)
                    ON CONFLICT (projeto) DO UPDATE SET
                        min_mecanica = EXCLUDED.min_mecanica,
                        min_instrumentacao = EXCLUDED.min_instrumentacao,
                        min_eletrica = EXCLUDED.min_eletrica,
                        min_total = EXCLUDED.min_total,
                        updated_at = EXCLUDED.updated_at
                """, (r["projeto"], r["min_mecanica"], r["min_instrumentacao"], r["min_eletrica"], r["min_total"], r["updated_at"]))
        pg_conn.commit()
        print(f"   ✅ {len(vagas_rows)} configurações de vagas migradas.")

        # 7. Migra p80_embarque_rt
        print("\n7️⃣ Migrando 'p80_embarque_rt'...")
        cur_sq.execute("SELECT projeto, nome, data_inicio, data_fim, numero_rt, updated_at FROM p80_embarque_rt")
        rt_rows = cur_sq.fetchall()
        with pg_conn.cursor() as cur:
            for r in rt_rows:
                cur.execute("""
                    INSERT INTO p80_embarque_rt (projeto, nome, data_inicio, data_fim, numero_rt, updated_at)
                    VALUES (%s, %s, %s, %s, %s, %s)
                    ON CONFLICT (projeto, nome, data_inicio, data_fim) DO UPDATE SET
                        numero_rt = EXCLUDED.numero_rt,
                        updated_at = EXCLUDED.updated_at
                """, (r["projeto"], r["nome"], r["data_inicio"], r["data_fim"], r["numero_rt"], r["updated_at"]))
        pg_conn.commit()
        print(f"   ✅ {len(rt_rows)} RTs de embarques migradas.")

        # Ajusta sequências restantes
        with pg_conn.cursor() as cur:
            cur.execute("SELECT setval(pg_get_serial_sequence('p80_membros', 'id'), coalesce(max(id), 1)) FROM p80_membros")
            cur.execute("SELECT setval(pg_get_serial_sequence('p80_escalas', 'id'), coalesce(max(id), 1)) FROM p80_escalas")
            cur.execute("SELECT setval(pg_get_serial_sequence('p80_embarque_rt', 'id'), coalesce(max(id), 1)) FROM p80_embarque_rt")
        pg_conn.commit()

        print("\n" + "=" * 70)
        print("🎉 MIGRAÇÃO CONCLUÍDA COM SUCESSO!")
        print("=" * 70)
        print("Todos os dados do SQLite agora estão salvos com persistência permanente no Supabase!")
        return True

    except Exception as err:
        pg_conn.rollback()
        print(f"\n❌ ERRO durante a migração: {err}")
        return False
    finally:
        sqlite_conn.close()
        pg_conn.close()


if __name__ == "__main__":
    cli_url = sys.argv[1] if len(sys.argv) > 1 else None
    success = run_migration(cli_url)
    sys.exit(0 if success else 1)
