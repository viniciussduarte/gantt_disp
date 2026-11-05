"""
Módulo para as abas de Escala de Projetos/Plataformas (P-80, P-82, P-83) no Streamlit.
Fornece:
- Persistência em banco de dados SQLite local (escala_p80.db)
- Suporte para múltiplas plataformas: P-80, P-82 e P-83
- Exibição por default dos membros nativos da equipe + membros externos escalados no mês
- Status: EMB (Embarcado), FOL (Folga), PMB, FÉRIAS (FÉR), BASE, TREINAMENTO (TREIN)
- Projeção automática de 1,5 dias de folga por dia de embarque (arredondado para cima com math.ceil)
- Perfis de acesso: Visualizador e Editor (com senha 'admin', mantendo o estado de edição 100% estável)
- Edição por clique direto no período ou seleção de período
- Limpeza de escalas robusta e ajustada
- Opção 'EXTERNO' com campo livre para digitar o nome do profissional
"""

import calendar
import math
import os
import sqlite3
from collections import defaultdict
from datetime import date, datetime, timedelta
from pathlib import Path
from typing import Dict, List, Optional, Tuple

import pandas as pd
import streamlit as st
import streamlit.components.v1 as components


# --- CONFIGURAÇÕES E CONSTANTES ---------------------------------------------
DB_PATH = Path(__file__).parent / "escala_p80.db"

PROJETOS_SUPORTADOS = ["P-80", "P-82", "P-83"]

DISCIPLINAS_PADRAO = ["Mecânica", "Instrumentação", "Elétrica", "Engenharia"]

MESES_PT = [
    "Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho",
    "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro"
]

DIAS_SEMANA_PT = ["Seg", "Ter", "Qua", "Qui", "Sex", "Sáb", "Dom"]

STATUS_MAP = {
    "EMB": {
        "nome": "Embarcado",
        "codigo": "EMB",
        "bg": "#075b9b",
        "color": "#FFFFFF",
        "borda": "#05487a",
        "desc": "Embarcado na unidade"
    },
    "FOL": {
        "nome": "Folga",
        "codigo": "FOL",
        "bg": "#64748b",
        "color": "#FFFFFF",
        "borda": "#475569",
        "desc": "Folga regulamentar"
    },
    "PMB": {
        "nome": "PMB",
        "codigo": "PMB",
        "bg": "#0891b2",
        "color": "#FFFFFF",
        "borda": "#0e7490",
        "desc": "PMB"
    },
    "FÉR": {
        "nome": "Férias",
        "codigo": "FÉR",
        "bg": "#d97706",
        "color": "#FFFFFF",
        "borda": "#b45309",
        "desc": "Período de férias regulamentares"
    },
    "BASE": {
        "nome": "Escritório / Base",
        "codigo": "BASE",
        "bg": "#7c3aed",
        "color": "#FFFFFF",
        "borda": "#6d28d9",
        "desc": "Trabalho em base ou escritório"
    },
    "TREIN": {
        "nome": "Treinamento",
        "codigo": "TREIN",
        "bg": "#059669",
        "color": "#FFFFFF",
        "borda": "#047857",
        "desc": "Treinamento / Certificação"
    },
    "CONF": {
        "nome": "Conflito",
        "codigo": "CONF",
        "bg": "#dc2626",
        "color": "#FFFFFF",
        "borda": "#b91c1c",
        "desc": "Conflito de agendamento detectado"
    },
    "LIC": {
        "nome": "Licença",
        "codigo": "LIC",
        "bg": "#9f1239",
        "color": "#FFFFFF",
        "borda": "#881337",
        "desc": "Licença"
    },
    "EMP": {
        "nome": "Emprestado",
        "codigo": "EMP",
        "bg": "#4338ca",
        "color": "#FFFFFF",
        "borda": "#312e81",
        "desc": "Emprestado para outra unidade"
    },
}

# Aliases para compatibilidade retroativa
STATUS_MAP["FÉRIAS"] = STATUS_MAP["FÉR"]
STATUS_MAP["TREINAMENTO"] = STATUS_MAP["TREIN"]
STATUS_MAP["E"] = STATUS_MAP["EMB"]
STATUS_MAP["CONFLITO"] = STATUS_MAP["CONF"]
STATUS_MAP["ESTALEIRO"] = STATUS_MAP["BASE"]
STATUS_MAP["LICENÇA"] = STATUS_MAP["LIC"]
STATUS_MAP["LICENCA"] = STATUS_MAP["LIC"]
STATUS_MAP["EMPRESTADO"] = STATUS_MAP["EMP"]


def is_mecanica(d: Optional[str]) -> bool:
    """Identifica se a disciplina é Mecânica independentemente de encoding ou acentuação."""
    u = str(d or "").upper()
    return "MEC" in u


def is_instrumentacao(d: Optional[str]) -> bool:
    """Identifica se a disciplina é Instrumentação independentemente de encoding ou acentuação."""
    u = str(d or "").upper()
    return "INST" in u


def is_eletrica(d: Optional[str]) -> bool:
    """Identifica se a disciplina é Elétrica independentemente de encoding ou acentuação."""
    u = str(d or "").upper()
    return "ELET" in u or "ELÉT" in u or ("EL" in u and "TRIC" in u)


def normalizar_status(st: Optional[str]) -> str:
    """Padroniza códigos de status para as chaves principais."""
    if not st:
        return "EMB"
    s = str(st).strip().upper()
    if s in ("E", "EMB", "EMBARCADO"):
        return "EMB"
    if s in ("FOL", "FOLGA"):
        return "FOL"
    if s in ("PMB", "PRE-EMBARQUE", "PRÉ-EMBARQUE"):
        return "PMB"
    if s in ("FER", "FÉR", "FÉRIAS", "FERIAS"):
        return "FÉR"
    if s in ("BASE", "ESCRITORIO", "ESCRITÓRIO", "ESTALEIRO"):
        return "BASE"
    if s in ("TREIN", "TREINAMENTO"):
        return "TREIN"
    if s in ("CONF", "CONFLITO"):
        return "CONF"
    if s in ("LIC", "LICENÇA", "LICENCA"):
        return "LIC"
    if s in ("EMP", "EMPRESTADO"):
        return "EMP"
    return s


def normalizar_projeto(p: Optional[str]) -> str:
    """Padroniza nomenclaturas como 'P80' -> 'P-80'."""
    if not p:
        return "P-80"
    p_limpo = str(p).strip().upper().replace(" ", "").replace("-", "")
    if "80" in p_limpo:
        return "P-80"
    if "82" in p_limpo:
        return "P-82"
    if "83" in p_limpo:
        return "P-83"
    if "BAC" in p_limpo:
        return "PBAC"
    return str(p).strip()


def get_current_user_name() -> str:
    """Retorna o identificador do usuário atual que realizou a alteração (usuário SO + perfil)."""
    os_user = os.environ.get("USERNAME") or os.environ.get("USER") or "Administrador"
    try:
        if st.session_state.get("escala_is_editor"):
            return f"{os_user} (Editor)"
    except Exception:
        pass
    return os_user


# --- GERENCIADOR DO BANCO DE DADOS ------------------------------------------
class EscalaP80DB:
    """Gerencia a persistência das Escalas de Projetos em SQLite."""

    @staticmethod
    def get_connection():
        conn = sqlite3.connect(DB_PATH)
        conn.row_factory = sqlite3.Row
        return conn

    @classmethod
    def init_db(cls):
        """Inicializa as tabelas com suporte a múltiplos projetos, sugestões e logs de auditoria."""
        with cls.get_connection() as conn:
            cursor = conn.cursor()
            cursor.execute("""
                CREATE TABLE IF NOT EXISTS p80_membros (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    nome TEXT UNIQUE NOT NULL,
                    disciplina TEXT NOT NULL,
                    funcao TEXT,
                    projeto TEXT DEFAULT 'P-80',
                    ativo INTEGER DEFAULT 1,
                    ordem INTEGER DEFAULT 0
                )
            """)
            cursor.execute("""
                CREATE TABLE IF NOT EXISTS p80_escalas (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
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

            # Tabela de Sugestões de Implementação / Melhorias
            cursor.execute("""
                CREATE TABLE IF NOT EXISTS p80_sugestoes (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
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

            # Tabela de Log de Auditoria & Histórico de Edições
            cursor.execute("""
                CREATE TABLE IF NOT EXISTS p80_audit_log (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    data_hora TEXT NOT NULL,
                    usuario TEXT NOT NULL,
                    acao TEXT NOT NULL,
                    projeto TEXT DEFAULT 'GERAL',
                    colaborador TEXT DEFAULT '—',
                    detalhes TEXT NOT NULL
                )
            """)
            cursor.execute("CREATE INDEX IF NOT EXISTS idx_audit_log_data ON p80_audit_log(data_hora DESC)")
            cursor.execute("CREATE INDEX IF NOT EXISTS idx_audit_log_proj ON p80_audit_log(projeto)")

            # Tabela de Configuração de Vagas Mínimas por Ativo (Mecânica, Instrumentação, Elétrica, Total)
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

            # Migração: coluna regime_folga em p80_membros ('14x21' ou '14x14')
            cursor.execute("PRAGMA table_info(p80_membros)")
            cols_membros = [row["name"] for row in cursor.fetchall()]
            if "regime_folga" not in cols_membros:
                cursor.execute("ALTER TABLE p80_membros ADD COLUMN regime_folga TEXT DEFAULT '14x21'")

            # Correção de encoding/acentos nas disciplinas
            cursor.execute("UPDATE p80_membros SET disciplina = 'Elétrica' WHERE disciplina LIKE '%TRIC%' OR disciplina LIKE '%ELET%'")
            cursor.execute("UPDATE p80_escalas SET disciplina = 'Elétrica' WHERE disciplina LIKE '%TRIC%' OR disciplina LIKE '%ELET%'")
            cursor.execute("UPDATE p80_membros SET disciplina = 'Mecânica' WHERE disciplina LIKE '%MEC%'")
            cursor.execute("UPDATE p80_escalas SET disciplina = 'Mecânica' WHERE disciplina LIKE '%MEC%'")
            cursor.execute("UPDATE p80_membros SET disciplina = 'Instrumentação' WHERE disciplina LIKE '%INST%'")
            cursor.execute("UPDATE p80_escalas SET disciplina = 'Instrumentação' WHERE disciplina LIKE '%INST%'")

            cursor.execute("UPDATE p80_escalas SET status = 'EMB' WHERE status = 'E'")
            # Regra estrita: Não puxa supervisores/coordenadores para escalas de plataforma
            cursor.execute("UPDATE p80_membros SET ativo = 0 WHERE UPPER(funcao) LIKE '%SUPERVIS%' OR UPPER(funcao) LIKE '%COORD%'")
            conn.commit()

    @classmethod
    def get_regime_colaborador(cls, nome: str) -> str:
        """Retorna o regime de folga configurado no cadastro do colaborador ('14x21' ou '14x14')."""
        with cls.get_connection() as conn:
            cursor = conn.cursor()
            cursor.execute("SELECT regime_folga FROM p80_membros WHERE UPPER(TRIM(nome)) = ?", (nome.strip().upper(),))
            row = cursor.fetchone()
            if row and row["regime_folga"]:
                reg = str(row["regime_folga"]).strip()
                if "14" in reg and "21" not in reg:
                    return "14x14"
        return "14x21"

    @classmethod
    def set_regime_colaborador(cls, nome: str, regime: str) -> bool:
        """Atualiza o regime de folga no cadastro do colaborador."""
        reg_final = "14x14" if ("14" in str(regime) and "21" not in str(regime)) else "14x21"
        with cls.get_connection() as conn:
            cursor = conn.cursor()
            cursor.execute("UPDATE p80_membros SET regime_folga = ? WHERE UPPER(TRIM(nome)) = ?", (reg_final, nome.strip().upper()))
            conn.commit()
            return cursor.rowcount > 0

    @classmethod
    def get_config_vagas(cls, projeto: str) -> Dict[str, int]:
        """Obtém os parâmetros de vagas mínimas da plataforma."""
        proj_norm = normalizar_projeto(projeto)
        with cls.get_connection() as conn:
            cursor = conn.cursor()
            cursor.execute("SELECT min_mecanica, min_instrumentacao, min_eletrica, min_total FROM p80_config_vagas WHERE projeto = ?", (proj_norm,))
            row = cursor.fetchone()
            if row:
                return {
                    "min_mecanica": int(row["min_mecanica"] or 0),
                    "min_instrumentacao": int(row["min_instrumentacao"] or 0),
                    "min_eletrica": int(row["min_eletrica"] or 0),
                    "min_total": int(row["min_total"] or 0)
                }
        return {"min_mecanica": 0, "min_instrumentacao": 0, "min_eletrica": 0, "min_total": 0}

    @classmethod
    def set_config_vagas(
        cls,
        projeto: str,
        min_mec: int = 0,
        min_inst: int = 0,
        min_elet: int = 0,
        min_total: int = 0,
        min_mecanica: Optional[int] = None,
        min_instrumentacao: Optional[int] = None,
        min_eletrica: Optional[int] = None
    ) -> bool:
        if min_mecanica is not None:
            min_mec = min_mecanica
        if min_instrumentacao is not None:
            min_inst = min_instrumentacao
        if min_eletrica is not None:
            min_elet = min_eletrica
        """Salva os parâmetros de vagas mínimas para a plataforma informada."""
        proj_norm = normalizar_projeto(projeto)
        with cls.get_connection() as conn:
            cursor = conn.cursor()
            cursor.execute("""
                INSERT INTO p80_config_vagas (projeto, min_mecanica, min_instrumentacao, min_eletrica, min_total, updated_at)
                VALUES (?, ?, ?, ?, ?, CURRENT_TIMESTAMP)
                ON CONFLICT(projeto) DO UPDATE SET
                    min_mecanica = excluded.min_mecanica,
                    min_instrumentacao = excluded.min_instrumentacao,
                    min_eletrica = excluded.min_eletrica,
                    min_total = excluded.min_total,
                    updated_at = CURRENT_TIMESTAMP
            """, (proj_norm, int(min_mec), int(min_inst), int(min_elet), int(min_total)))
            conn.commit()
            return True

    @classmethod
    def sync_membros_from_equipe(cls, equipe_df: Optional[pd.DataFrame]) -> int:
        """Importa membros da planilha de equipe do Excel caso ainda não existam (exclui supervisores das escalas)."""
        if equipe_df is None or equipe_df.empty:
            return 0

        disc_map = {
            "MEC": "Mecânica",
            "MECÂNICA": "Mecânica",
            "MECANICA": "Mecânica",
            "INST": "Instrumentação",
            "INSTRUMENTAÇÃO": "Instrumentação",
            "INSTRUMENTACAO": "Instrumentação",
            "ELET": "Elétrica",
            "ELÉTRICA": "Elétrica",
            "ELETRICA": "Elétrica",
            "COORD": "Engenharia",
            "ENG": "Engenharia",
            "ENGENHARIA": "Engenharia",
            "SUPERVISOR": "Engenharia",
            "LIDERANÇA": "Engenharia",
        }

        count = 0
        with cls.get_connection() as conn:
            cursor = conn.cursor()
            for _, row in equipe_df.iterrows():
                nome = str(row.get("Nome", "")).strip().upper()
                if not nome or nome == "NAN":
                    continue
                orig_disc = str(row.get("Disciplina", "")).strip().upper()
                disciplina = disc_map.get(orig_disc, "Engenharia")
                funcao = str(row.get("Função", "Colaborador")).strip().title()
                projeto_raw = str(row.get("Projeto", "P-80")).strip()
                projeto = normalizar_projeto(projeto_raw)

                # Regra: Não puxar supervisores para escalas de plataforma
                if "SUPERVIS" in funcao.upper() or "COORD" in funcao.upper():
                    continue

                cursor.execute("""
                    INSERT INTO p80_membros (nome, disciplina, funcao, projeto, ativo, regime_folga)
                    VALUES (?, ?, ?, ?, 1, '14x21')
                    ON CONFLICT(nome) DO UPDATE SET
                        disciplina = excluded.disciplina,
                        funcao = excluded.funcao,
                        projeto = excluded.projeto,
                        ativo = excluded.ativo
                """, (nome, disciplina, funcao, projeto))
                if cursor.rowcount > 0:
                    count += 1
            conn.commit()
        return count

    @classmethod
    def get_todos_membros(cls, apenas_ativos: bool = True, excluir_supervisores: bool = True) -> List[Dict]:
        """Retorna todos os membros cadastrados na base (por padrão, exclui supervisores para as escalas)."""
        with cls.get_connection() as conn:
            cursor = conn.cursor()
            query = "SELECT id, nome, disciplina, funcao, projeto, ativo, COALESCE(regime_folga, '14x21') as regime_folga FROM p80_membros WHERE 1=1"
            if apenas_ativos:
                query += " AND ativo = 1"
            if excluir_supervisores:
                query += " AND UPPER(funcao) NOT LIKE '%SUPERVIS%' AND UPPER(funcao) NOT LIKE '%COORD%'"
            query += " ORDER BY disciplina, nome"
            cursor.execute(query)
            return [dict(row) for row in cursor.fetchall()]

    @classmethod
    def get_membros_para_escala(
        cls,
        projeto: str,
        d_inicio_mes: str,
        d_fim_mes: str
    ) -> Tuple[List[Dict], int, int]:
        """
        Retorna a lista de profissionais a serem exibidos na grade daquele projeto:
        1. Membros nativos da equipe deste projeto (excluindo supervisores).
        2. Membros externos que tenham pelo menos 1 escala neste projeto dentro do mês (excluindo supervisores).
        """
        proj_norm = normalizar_projeto(projeto)
        todos = cls.get_todos_membros(apenas_ativos=True, excluir_supervisores=True)

        membros_nativos = [
            m for m in todos
            if normalizar_projeto(m.get("projeto")) == proj_norm
        ]

        with cls.get_connection() as conn:
            cursor = conn.cursor()
            cursor.execute("""
                SELECT DISTINCT nome FROM p80_escalas
                WHERE projeto = ? AND data >= ? AND data <= ?
            """, (proj_norm, d_inicio_mes, d_fim_mes))
            nomes_escalados = {row["nome"] for row in cursor.fetchall()}

        nomes_nativos = {m["nome"] for m in membros_nativos}
        nomes_externos = nomes_escalados - nomes_nativos

        todos_map = {m["nome"]: m for m in todos}
        membros_externos = []
        for nome_ext in nomes_externos:
            if nome_ext in todos_map:
                m_obj = dict(todos_map[nome_ext])
                if "SUPERVIS" in m_obj.get("funcao", "").upper() or "COORD" in m_obj.get("funcao", "").upper():
                    continue
                m_obj["is_externo"] = True
                membros_externos.append(m_obj)
            else:
                membros_externos.append({
                    "nome": nome_ext,
                    "disciplina": "Mecânica",
                    "funcao": "Externo",
                    "projeto": "EXTERNO",
                    "ativo": 1,
                    "is_externo": True
                })

        for m in membros_nativos:
            m["is_externo"] = False

        ordem_disc = {d: i for i, d in enumerate(DISCIPLINAS_PADRAO)}
        membros_combinados = membros_nativos + membros_externos
        membros_combinados.sort(
            key=lambda x: (ordem_disc.get(x["disciplina"], 99), x["nome"])
        )

        return membros_combinados, len(membros_nativos), len(membros_externos)

    @classmethod
    def get_escalas_periodo(
        cls,
        projeto: str,
        d_inicio: str,
        d_fim: str,
        incluir_emprestimos: bool = True
    ) -> Dict[Tuple[str, str], str]:
        """
        Retorna dicionário {(nome, data_iso): status} para o projeto e período informado.
        Se incluir_emprestimos=True, membros nativos desta plataforma que estejam escalados com EMB
        em outra plataforma offshore aparecem marcados como 'EMP' (Emprestado).
        """
        proj_norm = normalizar_projeto(projeto)
        result = {}
        with cls.get_connection() as conn:
            cursor = conn.cursor()
            cursor.execute("""
                SELECT nome, data, status FROM p80_escalas
                WHERE projeto = ? AND data >= ? AND data <= ?
            """, (proj_norm, d_inicio, d_fim))
            for row in cursor.fetchall():
                st_val = normalizar_status(row["status"])
                result[(row["nome"], row["data"])] = st_val

            if incluir_emprestimos and proj_norm != "TODOS":
                # Identifica membros nativos deste projeto
                cursor.execute("""
                    SELECT nome FROM p80_membros
                    WHERE ativo = 1 AND projeto = ?
                """, (proj_norm,))
                nativos = [row["nome"] for row in cursor.fetchall()]
                if nativos:
                    ph = ",".join(["?"] * len(nativos))
                    cursor.execute(f"""
                        SELECT nome, projeto, data, status FROM p80_escalas
                        WHERE projeto != ? AND status IN ('EMB', 'E') AND data >= ? AND data <= ?
                        AND nome IN ({ph})
                    """, [proj_norm, d_inicio, d_fim] + nativos)
                    for row in cursor.fetchall():
                        # Marca como EMP na plataforma nativa de origem
                        result[(row["nome"], row["data"])] = "EMP"

        return result

    @classmethod
    def get_periodos_agrupados(cls, projeto: str) -> Tuple[List[Dict], Dict[Tuple[str, str], Dict]]:
        """
        Retorna:
        1. Lista de períodos contínuos cadastrados no projeto:
           [{'nome', 'disciplina', 'projeto', 'inicio', 'fim', 'status', 'dias'}, ...]
        2. Dicionário {(nome, data_iso): periodo_dict} mapeando cada dia ao seu período.
        """
        with cls.get_connection() as conn:
            cursor = conn.cursor()
            if projeto == "TODOS":
                cursor.execute("""
                    SELECT nome, disciplina, projeto, data, status, COALESCE(observacao, '') as observacao
                    FROM p80_escalas
                    ORDER BY nome, data
                """)
            else:
                proj_norm = normalizar_projeto(projeto)
                cursor.execute("""
                    SELECT nome, disciplina, projeto, data, status, COALESCE(observacao, '') as observacao
                    FROM p80_escalas
                    WHERE projeto = ?
                    ORDER BY nome, data
                """, (proj_norm,))
            rows = cursor.fetchall()

        por_pessoa = defaultdict(list)
        for r in rows:
            por_pessoa[(r["nome"], r["disciplina"], r["projeto"])].append(
                (r["data"], normalizar_status(r["status"]), r["observacao"], None)
            )

        # Se for um projeto específico, detecta colaboradores nativos emprestados para outros projetos
        if projeto != "TODOS":
            proj_norm = normalizar_projeto(projeto)
            with cls.get_connection() as conn:
                cursor = conn.cursor()
                cursor.execute("SELECT nome, disciplina FROM p80_membros WHERE ativo = 1 AND projeto = ?", (proj_norm,))
                nativos = {row["nome"]: row["disciplina"] for row in cursor.fetchall()}
                if nativos:
                    ph = ",".join(["?"] * len(nativos))
                    cursor.execute(f"""
                        SELECT nome, disciplina, projeto, data, status, COALESCE(observacao, '') as observacao
                        FROM p80_escalas
                        WHERE projeto != ? AND status IN ('EMB', 'E') AND nome IN ({ph})
                        ORDER BY nome, data
                    """, [proj_norm] + list(nativos.keys()))
                    for r in cursor.fetchall():
                        disc_nat = nativos.get(r["nome"], r["disciplina"])
                        obs_emp = f"Emprestado para {r['projeto']}" + (f": {r['observacao']}" if r['observacao'] else "")
                        por_pessoa[(r["nome"], disc_nat, proj_norm)].append(
                            (r["data"], "EMP", obs_emp, r["projeto"])
                        )

        lista_periodos = []
        mapa_dia_periodo = {}

        for (nome, disc, proj), lista in por_pessoa.items():
            lista.sort(key=lambda x: x[0])
            p_ini, p_fim, p_st, p_obs, datas_bloco, p_dest = None, None, None, "", [], None
            for dt_str, st_code, obs_str, dest_proj in lista:
                dt = datetime.strptime(dt_str, "%Y-%m-%d").date()
                if p_ini is None:
                    p_ini, p_fim, p_st, p_obs, datas_bloco, p_dest = dt, dt, st_code, obs_str, [dt_str], dest_proj
                elif st_code == p_st and dt == p_fim + timedelta(days=1) and dest_proj == p_dest:
                    p_fim = dt
                    datas_bloco.append(dt_str)
                    if obs_str and not p_obs:
                        p_obs = obs_str
                else:
                    p_obj = {
                        "nome": nome, "disciplina": disc, "projeto": proj,
                        "inicio": p_ini, "fim": p_fim, "status": p_st,
                        "dias": (p_fim - p_ini).days + 1,
                        "observacao": p_obs,
                        "is_emprestado": (p_st == "EMP"),
                        "projeto_destino": p_dest
                    }
                    lista_periodos.append(p_obj)
                    for d_iso in datas_bloco:
                        mapa_dia_periodo[(nome, d_iso)] = p_obj
                    p_ini, p_fim, p_st, p_obs, datas_bloco, p_dest = dt, dt, st_code, obs_str, [dt_str], dest_proj

            if p_ini is not None:
                p_obj = {
                    "nome": nome, "disciplina": disc, "projeto": proj,
                    "inicio": p_ini, "fim": p_fim, "status": p_st,
                    "dias": (p_fim - p_ini).days + 1,
                    "observacao": p_obs,
                    "is_emprestado": (p_st == "EMP"),
                    "projeto_destino": p_dest
                }
                lista_periodos.append(p_obj)
                for d_iso in datas_bloco:
                    mapa_dia_periodo[(nome, d_iso)] = p_obj

        # Ordena períodos por data de início decrescente
        lista_periodos.sort(key=lambda x: (x["inicio"], x["nome"]), reverse=True)
        return lista_periodos, mapa_dia_periodo

    @classmethod
    def salvar_escala_periodo(
        cls,
        nome: str,
        disciplina: str,
        projeto: str,
        d_inicio: date,
        d_fim: date,
        status: str,
        observacao: str = "",
        regime_folga: Optional[str] = None,
        skip_log: bool = False
    ) -> Tuple[int, int]:
        """
        Salva escala para cada dia do intervalo [d_inicio, d_fim] no projeto informado.
        Regra obrigatória de negócio:
        - Se status == 'EMB':
          * Se regime_folga == '14x14': gera 1 dia de folga por dia de embarque (1x).
          * Se regime_folga == '14x21': gera 1,5 dias de folga por dia de embarque (ceil(1,5x)).
        - Se status == 'LIC':
          * Não gera projeção de folga (apenas marca o período).
        Retorna (dias_salvos, dias_folga_projetados).
        """
        nome = nome.strip().upper()
        proj_norm = normalizar_projeto(projeto)
        st_final = normalizar_status(status)

        if d_inicio > d_fim:
            d_inicio, d_fim = d_fim, d_inicio

        total_dias = (d_fim - d_inicio).days + 1
        dias_salvos = 0
        dias_folga_projetados = 0

        # Se informado regime no lançamento, persiste no cadastro do colaborador
        if regime_folga:
            cls.set_regime_colaborador(nome, regime_folga)

        regime_atual = regime_folga or cls.get_regime_colaborador(nome)
        is_14x14 = ("14" in str(regime_atual) and "21" not in str(regime_atual))

        with cls.get_connection() as conn:
            cursor = conn.cursor()
            curr = d_inicio
            while curr <= d_fim:
                iso_date = curr.strftime("%Y-%m-%d")
                cursor.execute("""
                    INSERT INTO p80_escalas (nome, disciplina, projeto, data, status, observacao, updated_at)
                    VALUES (?, ?, ?, ?, ?, ?, CURRENT_TIMESTAMP)
                    ON CONFLICT(nome, projeto, data) DO UPDATE SET
                        disciplina = excluded.disciplina,
                        status = excluded.status,
                        observacao = excluded.observacao,
                        updated_at = CURRENT_TIMESTAMP
                """, (nome, disciplina, proj_norm, iso_date, st_final, observacao))
                curr += timedelta(days=1)
                dias_salvos += 1

            if st_final == "EMB":
                if is_14x14:
                    qtd_folga = total_dias
                    obs_folga_auto = "Projeção automática de folga (14x14 - 1x)"
                else:
                    qtd_folga = math.ceil(total_dias * 1.5)
                    obs_folga_auto = "Projeção automática de folga (14x21 - 1,5x)"

                curr_folga = d_fim + timedelta(days=1)
                for _ in range(qtd_folga):
                    iso_folga = curr_folga.strftime("%Y-%m-%d")
                    cursor.execute("""
                        INSERT INTO p80_escalas (nome, disciplina, projeto, data, status, observacao, updated_at)
                        VALUES (?, ?, ?, ?, 'FOL', ?, CURRENT_TIMESTAMP)
                        ON CONFLICT(nome, projeto, data) DO UPDATE SET
                            disciplina = excluded.disciplina,
                            status = excluded.status,
                            observacao = excluded.observacao,
                            updated_at = CURRENT_TIMESTAMP
                    """, (nome, disciplina, proj_norm, iso_folga, obs_folga_auto))
                    curr_folga += timedelta(days=1)
                    dias_folga_projetados += 1

            conn.commit()

        if not skip_log:
            cls.registrar_log(
                acao="INSERÇÃO",
                detalhes=f"Lançou período de '{st_final}' de {d_inicio:%d/%m/%Y} a {d_fim:%d/%m/%Y} ({total_dias} dias)." + (f" Projeção automática gerada: {dias_folga_projetados} dias de folga (FOL - Regime {regime_atual})." if dias_folga_projetados > 0 else "") + (f" Obs: {observacao}" if observacao else ""),
                projeto=proj_norm,
                colaborador=nome
            )

        return dias_salvos, dias_folga_projetados

    @classmethod
    def salvar_escala_ciclica(
        cls,
        nome: str,
        disciplina: str,
        projeto: str,
        d_inicio: date,
        qtd_ciclos: int = 7,
        observacao: str = "",
        regime: str = "14x21"
    ) -> Tuple[int, int]:
        """
        Gera QTD_CICLOS consecutivos no regime offshore (14x21 ou 14x14) a partir de d_inicio:
        - 14x21: Cada ciclo contém 14 dias de embarque (EMB) e 21 dias de folga (FOL) - total 35 dias.
        - 14x14: Cada ciclo contém 14 dias de embarque (EMB) e 14 dias de folga (FOL) - total 28 dias.
        Retorna (total_dias_emb, total_dias_fol).
        """
        proj_norm = normalizar_projeto(projeto)
        nome = nome.strip().upper()
        if qtd_ciclos < 1:
            qtd_ciclos = 1
        if qtd_ciclos > 36:
            qtd_ciclos = 36

        is_14x14 = ("14" in str(regime) and "21" not in str(regime))
        reg_final = "14x14" if is_14x14 else "14x21"
        cls.set_regime_colaborador(nome, reg_final)

        dias_emb_c = 14
        dias_fol_c = 14 if is_14x14 else 21
        dias_ciclo = dias_emb_c + dias_fol_c

        total_emb = 0
        total_fol = 0

        with cls.get_connection() as conn:
            cursor = conn.cursor()
            for c_idx in range(qtd_ciclos):
                c_num = c_idx + 1
                c_offset_days = c_idx * dias_ciclo
                emb_ini = d_inicio + timedelta(days=c_offset_days)
                emb_fim = emb_ini + timedelta(days=dias_emb_c - 1)  # 14 dias
                fol_ini = emb_fim + timedelta(days=1)
                fol_fim = fol_ini + timedelta(days=dias_fol_c - 1)  # 14 ou 21 dias

                obs_emb = observacao if observacao else f"Escala {reg_final} (Ciclo {c_num}/{qtd_ciclos})"
                obs_fol = f"Folga {reg_final} (Ciclo {c_num}/{qtd_ciclos} - {dias_fol_c} dias)"

                # 1. 14 dias de embarque (EMB)
                curr = emb_ini
                while curr <= emb_fim:
                    iso_dt = curr.strftime("%Y-%m-%d")
                    cursor.execute("""
                        INSERT INTO p80_escalas (nome, disciplina, projeto, data, status, observacao, updated_at)
                        VALUES (?, ?, ?, ?, 'EMB', ?, CURRENT_TIMESTAMP)
                        ON CONFLICT(nome, projeto, data) DO UPDATE SET
                            disciplina = excluded.disciplina,
                            status = 'EMB',
                            observacao = excluded.observacao,
                            updated_at = CURRENT_TIMESTAMP
                    """, (nome, disciplina, proj_norm, iso_dt, obs_emb))
                    curr += timedelta(days=1)
                    total_emb += 1

                # 2. Folga (FOL)
                curr = fol_ini
                while curr <= fol_fim:
                    iso_dt = curr.strftime("%Y-%m-%d")
                    cursor.execute("""
                        INSERT INTO p80_escalas (nome, disciplina, projeto, data, status, observacao, updated_at)
                        VALUES (?, ?, ?, ?, 'FOL', ?, CURRENT_TIMESTAMP)
                        ON CONFLICT(nome, projeto, data) DO UPDATE SET
                            disciplina = excluded.disciplina,
                            status = 'FOL',
                            observacao = excluded.observacao,
                            updated_at = CURRENT_TIMESTAMP
                    """, (nome, disciplina, proj_norm, iso_dt, obs_fol))
                    curr += timedelta(days=1)
                    total_fol += 1

            conn.commit()

        cls.registrar_log(
            acao="INSERÇÃO",
            detalhes=f"Alocou escala contínua {reg_final} ({qtd_ciclos} ciclos) a partir de {d_inicio:%d/%m/%Y} (Total: {total_emb} dias EMB e {total_fol} dias FOL)." + (f" Obs: {observacao}" if observacao else ""),
            projeto=proj_norm,
            colaborador=nome
        )
        return total_emb, total_fol

    @classmethod
    def salvar_escala_ciclica_14x21(
        cls,
        nome: str,
        disciplina: str,
        projeto: str,
        d_inicio: date,
        qtd_ciclos: int = 7,
        observacao: str = "",
        regime: str = "14x21"
    ) -> Tuple[int, int]:
        return cls.salvar_escala_ciclica(nome, disciplina, projeto, d_inicio, qtd_ciclos, observacao, regime)

    @classmethod
    def atualizar_periodo(
        cls,
        nome: str,
        disciplina: str,
        projeto: str,
        d_ini_antigo: date,
        d_fim_antigo: date,
        d_ini_novo: date,
        d_fim_novo: date,
        status_novo: str,
        observacao: str = "",
        regime_folga: Optional[str] = None
    ) -> Tuple[int, int]:
        """
        Atualiza um período:
        Se o período antigo ou novo for EMB (Embarque):
        - Limpa o período antigo de embarque.
        - Limpa as folgas antigas subsequentes associadas ao embarque anterior.
        - Se a nova data de término for encurtada, limpa qualquer folga remanescente.
        - Grava o novo período com recálculo automático de folga a partir de d_fim_novo + 1 (respeitando 14x21 ou 14x14).
        """
        proj_norm = normalizar_projeto(projeto)
        nome = nome.strip().upper()
        st_novo_norm = normalizar_status(status_novo)

        d_ini_ant_str = d_ini_antigo.strftime("%Y-%m-%d")
        d_fim_ant_str = d_fim_antigo.strftime("%Y-%m-%d")

        with cls.get_connection() as conn:
            cursor = conn.cursor()

            # 1. Verifica se no intervalo antigo havia status EMB
            cursor.execute("""
                SELECT COUNT(*) FROM p80_escalas
                WHERE UPPER(TRIM(nome)) = ? AND UPPER(TRIM(projeto)) = ? AND data >= ? AND data <= ? AND status = 'EMB'
            """, (nome, proj_norm, d_ini_ant_str, d_fim_ant_str))
            tinha_emb_antigo = (cursor.fetchone()[0] > 0)

            # 2. Remove os dias antigos do período
            cursor.execute("""
                DELETE FROM p80_escalas
                WHERE UPPER(TRIM(nome)) = ? AND UPPER(TRIM(projeto)) = ? AND data >= ? AND data <= ?
            """, (nome, proj_norm, d_ini_ant_str, d_fim_ant_str))

            # 3. Se o período antigo continha EMB, remove folgas subsequentes associadas
            if tinha_emb_antigo:
                total_dias_ant = (d_fim_antigo - d_ini_antigo).days + 1
                qtd_folga_ant = math.ceil(total_dias_ant * 1.5) + 2
                f_ini_ant = (d_fim_antigo + timedelta(days=1)).strftime("%Y-%m-%d")
                f_fim_ant = (d_fim_antigo + timedelta(days=qtd_folga_ant)).strftime("%Y-%m-%d")

                cursor.execute("""
                    DELETE FROM p80_escalas
                    WHERE UPPER(TRIM(nome)) = ? AND UPPER(TRIM(projeto)) = ? AND data >= ? AND data <= ? AND status = 'FOL'
                """, (nome, proj_norm, f_ini_ant, f_fim_ant))

                if d_fim_novo < d_fim_antigo:
                    f_inter_ini = (d_fim_novo + timedelta(days=1)).strftime("%Y-%m-%d")
                    cursor.execute("""
                        DELETE FROM p80_escalas
                        WHERE UPPER(TRIM(nome)) = ? AND UPPER(TRIM(projeto)) = ? AND data >= ? AND data <= ? AND status = 'FOL'
                    """, (nome, proj_norm, f_inter_ini, f_fim_ant))

            conn.commit()

        # 4. Salva o novo período com projeção de folga apropriada
        res = cls.salvar_escala_periodo(
            nome=nome,
            disciplina=disciplina,
            projeto=proj_norm,
            d_inicio=d_ini_novo,
            d_fim=d_fim_novo,
            status=st_novo_norm,
            observacao=observacao,
            regime_folga=regime_folga,
            skip_log=True
        )

        cls.registrar_log(
            acao="EDIÇÃO",
            detalhes=f"Atualizou período: de {d_ini_antigo:%d/%m/%Y} a {d_fim_antigo:%d/%m/%Y} para {d_ini_novo:%d/%m/%Y} a {d_fim_novo:%d/%m/%Y} ('{st_novo_norm}'). Folgas recalculadas: {res[1]} dias." + (f" Obs: {observacao}" if observacao else ""),
            projeto=proj_norm,
            colaborador=nome
        )

        return res

    @classmethod
    def auto_permutar_folga(
        cls,
        nome: str,
        disciplina: str,
        projeto: str,
        origem_ini: date,
        origem_fim: date,
        destino_ini: date,
        observacao: str = ""
    ) -> Tuple[int, date]:
        """
        Executa a Auto-Permuta de Folga:
        - A pessoa troca seu período de folga [origem_ini, origem_fim] para usufruir
          a mesma quantidade contínua de dias a partir de destino_ini.
        - Regra de Negócio: destino_ini deve ser posterior a origem_ini e no máximo 180 dias após origem_fim!
        - Remove os dias de folga do intervalo de origem.
        - Grava os N dias contínuos de FOL no intervalo de destino [destino_ini, destino_fim].
        Retorna (qtd_dias, destino_fim).
        """
        proj_norm = normalizar_projeto(projeto)
        nome = nome.strip().upper()

        if origem_ini > origem_fim:
            origem_ini, origem_fim = origem_fim, origem_ini

        qtd_dias = (origem_fim - origem_ini).days + 1
        distancia = (destino_ini - origem_fim).days

        if destino_ini < origem_ini:
            raise ValueError("A nova data de início da folga deve ser posterior à data original.")
        if distancia > 180:
            raise ValueError(f"O limite máximo para a permuta é de até 180 dias (selecionado: {distancia} dias).")

        destino_fim = destino_ini + timedelta(days=qtd_dias - 1)

        d_orig_ini_str = origem_ini.strftime("%Y-%m-%d")
        d_orig_fim_str = origem_fim.strftime("%Y-%m-%d")

        obs_padrao = observacao if observacao else f"Auto-Permuta de folga (origem: {origem_ini:%d/%m/%Y} a {origem_fim:%d/%m/%Y})"

        with cls.get_connection() as conn:
            cursor = conn.cursor()

            # 1. Remove as folgas do período de origem
            cursor.execute("""
                DELETE FROM p80_escalas
                WHERE nome = ? AND projeto = ? AND data >= ? AND data <= ? AND status = 'FOL'
            """, (nome, proj_norm, d_orig_ini_str, d_orig_fim_str))

            # 2. Grava a nova folga contínua nas datas de destino
            curr = destino_ini
            while curr <= destino_fim:
                iso_dest = curr.strftime("%Y-%m-%d")
                cursor.execute("""
                    INSERT INTO p80_escalas (nome, disciplina, projeto, data, status, observacao, updated_at)
                    VALUES (?, ?, ?, ?, 'FOL', ?, CURRENT_TIMESTAMP)
                    ON CONFLICT(nome, projeto, data) DO UPDATE SET
                        disciplina = excluded.disciplina,
                        status = excluded.status,
                        observacao = excluded.observacao,
                        updated_at = CURRENT_TIMESTAMP
                """, (nome, disciplina, proj_norm, iso_dest, obs_padrao))
                curr += timedelta(days=1)

            conn.commit()

        cls.registrar_log(
            acao="AUTO-PERMUTA",
            detalhes=f"Auto-permuta de {qtd_dias} dias de folga: de {origem_ini:%d/%m/%Y} a {origem_fim:%d/%m/%Y} transferido para {destino_ini:%d/%m/%Y} a {destino_fim:%d/%m/%Y}." + (f" Obs: {obs_padrao}"),
            projeto=proj_norm,
            colaborador=nome
        )

        return qtd_dias, destino_fim

    @classmethod
    def excluir_periodo(
        cls,
        nome: str,
        projeto: str,
        d_inicio: date,
        d_fim: date,
        remover_folgas_subsequentes: bool = True
    ) -> int:
        """
        Exclui as escalas de um colaborador no intervalo [d_inicio, d_fim].
        Se for embarque e remover_folgas_subsequentes=True, remove também as folgas associadas.
        """
        proj_norm = normalizar_projeto(projeto)
        nome = nome.strip().upper()
        if d_inicio > d_fim:
            d_inicio, d_fim = d_fim, d_inicio
        d_ini_str = d_inicio.strftime("%Y-%m-%d")
        d_fim_str = d_fim.strftime("%Y-%m-%d")
        total_dias = (d_fim - d_inicio).days + 1
        removidos = 0

        with cls.get_connection() as conn:
            cursor = conn.cursor()

            # Verifica se no intervalo havia status EMB
            cursor.execute("""
                SELECT COUNT(*) FROM p80_escalas
                WHERE UPPER(TRIM(nome)) = ? AND UPPER(TRIM(projeto)) = ? AND data >= ? AND data <= ? AND status = 'EMB'
            """, (nome, proj_norm, d_ini_str, d_fim_str))
            tinha_emb = (cursor.fetchone()[0] > 0)

            cursor.execute("""
                DELETE FROM p80_escalas
                WHERE UPPER(TRIM(nome)) = ? AND UPPER(TRIM(projeto)) = ? AND data >= ? AND data <= ?
            """, (nome, proj_norm, d_ini_str, d_fim_str))
            removidos += cursor.rowcount

            if remover_folgas_subsequentes and tinha_emb:
                qtd_folga = math.ceil(total_dias * 1.5)
                f_ini_str = (d_fim + timedelta(days=1)).strftime("%Y-%m-%d")
                f_fim_str = (d_fim + timedelta(days=qtd_folga + 2)).strftime("%Y-%m-%d")
                cursor.execute("""
                    DELETE FROM p80_escalas
                    WHERE UPPER(TRIM(nome)) = ? AND UPPER(TRIM(projeto)) = ? AND data >= ? AND data <= ? AND status = 'FOL'
                """, (nome, proj_norm, f_ini_str, f_fim_str))
                removidos += cursor.rowcount

            conn.commit()

        cls.registrar_log(
            acao="EXCLUSÃO",
            detalhes=f"Excluiu período de {d_inicio:%d/%m/%Y} a {d_fim:%d/%m/%Y} ({removidos} registros removidos, incluindo folgas associadas).",
            projeto=proj_norm,
            colaborador=nome
        )

        return removidos

    @classmethod
    def limpar_escalas(
        cls,
        projeto: str,
        d_inicio: str,
        d_fim: str,
        nome: Optional[str] = None
    ) -> int:
        """
        Limpa escalas em um período e projeto. Se nome for informado e diferente de 'TODOS',
        limpa apenas daquele colaborador. Se nome for None ou 'TODOS', limpa de toda a plataforma no período.
        """
        proj_norm = normalizar_projeto(projeto)
        with cls.get_connection() as conn:
            cursor = conn.cursor()
            if nome and nome.strip().upper() != "TODOS":
                cursor.execute("""
                    DELETE FROM p80_escalas
                    WHERE projeto = ? AND nome = ? AND data >= ? AND data <= ?
                """, (proj_norm, nome.strip().upper(), d_inicio, d_fim))
            else:
                cursor.execute("""
                    DELETE FROM p80_escalas
                    WHERE projeto = ? AND data >= ? AND data <= ?
                """, (proj_norm, d_inicio, d_fim))
            removidos = cursor.rowcount
            conn.commit()

        alvo = f"colaborador {nome.strip().upper()}" if (nome and nome.strip().upper() != "TODOS") else "toda a unidade"
        cls.registrar_log(
            acao="LIMPEZA",
            detalhes=f"Limpeza de escalas em {proj_norm} para {alvo} no período de {d_inicio} a {d_fim} ({removidos} registros apagados).",
            projeto=proj_norm,
            colaborador=nome.strip().upper() if (nome and nome.strip().upper() != "TODOS") else "TODOS"
        )

        return removidos

    @classmethod
    def add_membro(cls, nome: str, disciplina: str, funcao: str = "Colaborador", projeto: str = "P-80") -> bool:
        """Adiciona ou reativa um membro na base."""
        nome = nome.strip().upper()
        if not nome or not disciplina:
            return False
        proj_norm = normalizar_projeto(projeto)
        with cls.get_connection() as conn:
            cursor = conn.cursor()
            cursor.execute("""
                INSERT INTO p80_membros (nome, disciplina, funcao, projeto, ativo)
                VALUES (?, ?, ?, ?, 1)
                ON CONFLICT(nome) DO UPDATE SET
                    disciplina = excluded.disciplina,
                    funcao = excluded.funcao,
                    projeto = excluded.projeto,
                    ativo = 1
            """, (nome, disciplina, funcao, proj_norm))
            conn.commit()

        cls.registrar_log(
            acao="EQUIPE",
            detalhes=f"Cadastrou/reativou colaborador {nome} (Disciplina: {disciplina}, Função: {funcao}) na equipe {proj_norm}.",
            projeto=proj_norm,
            colaborador=nome
        )
        return True

    @classmethod
    def delete_membro(cls, nome: str) -> bool:
        """Desativa um membro."""
        with cls.get_connection() as conn:
            cursor = conn.cursor()
            cursor.execute("UPDATE p80_membros SET ativo = 0 WHERE nome = ?", (nome.strip().upper(),))
            conn.commit()

        cls.registrar_log(
            acao="EQUIPE",
            detalhes=f"Desativou colaborador {nome.strip().upper()} da equipe.",
            projeto="GERAL",
            colaborador=nome.strip().upper()
        )
        return True

    @classmethod
    def registrar_log(
        cls,
        acao: str,
        detalhes: str,
        projeto: str = "GERAL",
        colaborador: str = "—",
        usuario: Optional[str] = None
    ) -> None:
        """Registra uma ação no histórico de auditoria do sistema."""
        try:
            if not usuario:
                usuario = get_current_user_name()
            proj_norm = normalizar_projeto(projeto) if projeto != "GERAL" else "GERAL"
            agora = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            with cls.get_connection() as conn:
                cursor = conn.cursor()
                cursor.execute("""
                    INSERT INTO p80_audit_log (data_hora, usuario, acao, projeto, colaborador, detalhes)
                    VALUES (?, ?, ?, ?, ?, ?)
                """, (agora, usuario, acao.upper(), proj_norm, colaborador.strip().upper() if colaborador else "—", detalhes))
                conn.commit()
        except Exception as e:
            print(f"Erro ao registrar audit log: {e}")

    @classmethod
    def get_audit_logs(
        cls,
        limite: int = 500,
        filtro_acao: Optional[str] = None,
        filtro_projeto: Optional[str] = None,
        termo_busca: Optional[str] = None
    ) -> List[Dict]:
        """Retorna os registros de audit log com filtros opcionais."""
        with cls.get_connection() as conn:
            cursor = conn.cursor()
            query = "SELECT id, data_hora, usuario, acao, projeto, colaborador, detalhes FROM p80_audit_log WHERE 1=1"
            params = []
            if filtro_acao and filtro_acao != "Todas":
                query += " AND acao = ?"
                params.append(filtro_acao.upper())
            if filtro_projeto and filtro_projeto != "Todos":
                query += " AND projeto = ?"
                params.append(filtro_projeto.upper())
            if termo_busca and termo_busca.strip():
                t = f"%{termo_busca.strip()}%"
                query += " AND (detalhes LIKE ? OR colaborador LIKE ? OR usuario LIKE ?)"
                params.extend([t, t, t])
            query += " ORDER BY data_hora DESC, id DESC"
            if limite and limite > 0:
                query += f" LIMIT {int(limite)}"
            cursor.execute(query, params)
            return [dict(row) for row in cursor.fetchall()]

    @classmethod
    def criar_sugestao(
        cls,
        autor: str,
        is_anonimo: bool,
        titulo: str,
        descricao: str
    ) -> int:
        """Cria uma nova sugestão com status 'Pendente'."""
        autor_limpo = "Anônimo" if is_anonimo else (autor.strip() if autor and autor.strip() else "Anônimo")
        titulo_limpo = titulo.strip()
        desc_limpa = descricao.strip()
        agora = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

        with cls.get_connection() as conn:
            cursor = conn.cursor()
            cursor.execute("""
                INSERT INTO p80_sugestoes (autor, is_anonimo, titulo, descricao, prioridade, status, created_at, updated_at)
                VALUES (?, ?, ?, ?, 'Média', 'Pendente', ?, ?)
            """, (autor_limpo, 1 if is_anonimo else 0, titulo_limpo, desc_limpa, agora, agora))
            sug_id = cursor.lastrowid
            conn.commit()

        cls.registrar_log(
            acao="SUGESTÃO",
            detalhes=f"Nova sugestão #{sug_id} enviada por {autor_limpo}: '{titulo_limpo}'",
            projeto="GERAL",
            colaborador=autor_limpo
        )
        return sug_id

    @classmethod
    def get_sugestoes(cls, status_filtro: Optional[str] = None) -> List[Dict]:
        """
        Retorna sugestões cadastradas.
        Se status_filtro == 'Aprovada', ordena rigorosamente pela prioridade:
        Alta (1) -> Média (2) -> Baixa (3) e depois pela data de envio.
        """
        with cls.get_connection() as conn:
            cursor = conn.cursor()
            if status_filtro:
                if status_filtro == "Aprovada":
                    query = """
                        SELECT id, autor, is_anonimo, titulo, descricao, prioridade, status, created_at, updated_at, implementada_at
                        FROM p80_sugestoes
                        WHERE status = ?
                        ORDER BY
                            CASE prioridade
                                WHEN 'Alta' THEN 1
                                WHEN 'Média' THEN 2
                                WHEN 'Baixa' THEN 3
                                ELSE 4
                            END ASC,
                            created_at ASC
                    """
                elif status_filtro == "Implementada":
                    query = """
                        SELECT id, autor, is_anonimo, titulo, descricao, prioridade, status, created_at, updated_at, implementada_at
                        FROM p80_sugestoes
                        WHERE status = ?
                        ORDER BY implementada_at DESC, updated_at DESC
                    """
                else:
                    query = """
                        SELECT id, autor, is_anonimo, titulo, descricao, prioridade, status, created_at, updated_at, implementada_at
                        FROM p80_sugestoes
                        WHERE status = ?
                        ORDER BY created_at DESC
                    """
                cursor.execute(query, (status_filtro,))
            else:
                cursor.execute("""
                    SELECT id, autor, is_anonimo, titulo, descricao, prioridade, status, created_at, updated_at, implementada_at
                    FROM p80_sugestoes
                    ORDER BY created_at DESC
                """)
            return [dict(row) for row in cursor.fetchall()]

    @classmethod
    def avaliar_sugestao(cls, id_sugestao: int, prioridade: str, status: str) -> bool:
        """Aprova ou rejeita uma sugestão e define sua prioridade."""
        prio_norm = prioridade.strip().title()
        if prio_norm not in ("Alta", "Média", "Baixa"):
            prio_norm = "Média"
        st_norm = status.strip().title()
        agora = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

        with cls.get_connection() as conn:
            cursor = conn.cursor()
            cursor.execute("SELECT titulo, autor FROM p80_sugestoes WHERE id = ?", (id_sugestao,))
            row = cursor.fetchone()
            titulo = row["titulo"] if row else f"#{id_sugestao}"
            autor = row["autor"] if row else "—"

            cursor.execute("""
                UPDATE p80_sugestoes
                SET prioridade = ?, status = ?, updated_at = ?
                WHERE id = ?
            """, (prio_norm, st_norm, agora, id_sugestao))
            conn.commit()

        cls.registrar_log(
            acao="SUGESTÃO",
            detalhes=f"Sugestão #{id_sugestao} ('{titulo}') avaliada: status '{st_norm}' com prioridade '{prio_norm}'.",
            projeto="GERAL",
            colaborador=autor
        )
        return True

    @classmethod
    def marcar_sugestao_implementada(cls, id_sugestao: int) -> bool:
        """Marca uma sugestão aprovada como Implementada."""
        agora = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        with cls.get_connection() as conn:
            cursor = conn.cursor()
            cursor.execute("SELECT titulo, autor FROM p80_sugestoes WHERE id = ?", (id_sugestao,))
            row = cursor.fetchone()
            titulo = row["titulo"] if row else f"#{id_sugestao}"
            autor = row["autor"] if row else "—"

            cursor.execute("""
                UPDATE p80_sugestoes
                SET status = 'Implementada', implementada_at = ?, updated_at = ?
                WHERE id = ?
            """, (agora, agora, id_sugestao))
            conn.commit()

        cls.registrar_log(
            acao="SUGESTÃO",
            detalhes=f"Sugestão #{id_sugestao} ('{titulo}') foi marcada como IMPLEMENTADA.",
            projeto="GERAL",
            colaborador=autor
        )
        return True


# --- INTEGRAÇÃO COM DADOS DO CRONOGRAMA & DETECÇÃO DE CONFLITOS (v0.2) -------
def integrar_escalas_com_cronograma(
    escalas_map: Dict[Tuple[str, str], str],
    mapa_dia_periodo: Dict[Tuple[str, str], Dict],
    membros: List[Dict],
    cronograma_df: Optional[pd.DataFrame],
    proj_norm: str,
    ano: int,
    mes: int,
    dt_inicio_custom: Optional[date] = None,
    dt_fim_custom: Optional[date] = None
) -> Tuple[Dict[Tuple[str, str], str], Dict[Tuple[str, str], Dict], List[Dict], List[Dict], int]:
    """
    Integra os eventos do cronograma externo à grade mensal ou contínua da plataforma:
    - Tipos importados:
      * 'Estaleiro' -> Mapeado para status 'BASE' com observação 'Estaleiro (Cronograma)'
      * 'Férias' / 'Ferias' -> Mapeado para status 'FÉR' com observação 'Férias (Cronograma)'
      * 'Folga' -> Mapeado para status 'FOL' com observação 'Folga (Cronograma)'
    - Regra: NÃO calcula 1,5x de folga para estes eventos do cronograma!
    - Detecção de Conflitos:
      Se em um mesmo dia o colaborador já tiver escala lançada na plataforma (ex: EMB em P-80)
      e houver uma atividade divergente no cronograma (ex: Férias ou Estaleiro),
      a célula é marcada com status 'CONF' (Conflito, exibido em vermelho com a palavra CONF).
    Retorna (merged_escalas_map, merged_mapa_dia_periodo, lista_conflitos, membros_atualizados, qtd_externos).
    """
    merged_map = dict(escalas_map)
    merged_periodos = dict(mapa_dia_periodo)
    lista_conflitos = []
    membros_atualizados = list(membros)
    nomes_atuais = {m["nome"] for m in membros_atualizados}

    if cronograma_df is None or cronograma_df.empty:
        qtd_ext = sum(1 for m in membros_atualizados if m.get("is_externo"))
        return merged_map, merged_periodos, lista_conflitos, membros_atualizados, qtd_ext

    if dt_inicio_custom is not None and dt_fim_custom is not None:
        dt_ini_mes = dt_inicio_custom
        dt_fim_mes = dt_fim_custom
    else:
        _, dias_no_mes = calendar.monthrange(ano, mes)
        dt_ini_mes = date(ano, mes, 1)
        dt_fim_mes = date(ano, mes, dias_no_mes)

    # Filtra eventos que não sejam da própria escala da plataforma no SQLite
    df_crono = cronograma_df.copy()
    if "Detalhamento" in df_crono.columns:
        df_crono = df_crono[~df_crono["Detalhamento"].fillna("").astype(str).str.startswith("Escala ")].copy()

    if "Tipo" not in df_crono.columns:
        qtd_ext = sum(1 for m in membros_atualizados if m.get("is_externo"))
        return merged_map, merged_periodos, lista_conflitos, membros_atualizados, qtd_ext

    df_crono["_tipo_norm"] = df_crono["Tipo"].astype(str).str.strip().str.title()
    tipos_validos = {"Estaleiro", "Férias", "Ferias", "Folga"}
    df_crono = df_crono[df_crono["_tipo_norm"].isin(tipos_validos)].copy()

    if df_crono.empty:
        qtd_ext = sum(1 for m in membros_atualizados if m.get("is_externo"))
        return merged_map, merged_periodos, lista_conflitos, membros_atualizados, qtd_ext

    df_crono["_ini"] = pd.to_datetime(df_crono["Início"], errors="coerce")
    df_crono["_fim"] = pd.to_datetime(df_crono["Término"], errors="coerce")
    df_crono = df_crono.dropna(subset=["_ini", "_fim", "Nome"])

    ts_ini_mes = pd.Timestamp(dt_ini_mes)
    ts_fim_mes = pd.Timestamp(dt_fim_mes)
    df_crono = df_crono[(df_crono["_ini"] <= ts_fim_mes) & (df_crono["_fim"] >= ts_ini_mes)]

    for _, row in df_crono.iterrows():
        nome = str(row["Nome"]).strip().upper()
        func_row = str(row.get("Função", "")).strip().upper()
        if "SUPERVIS" in func_row or "COORD" in func_row:
            continue

        tipo_raw = str(row["_tipo_norm"]).strip()
        disc = str(row.get("Disciplina", "Mecânica")).strip()
        if not disc or disc == "nan" or disc == "—":
            disc = "Mecânica"

        row_plat = str(row.get("Plataforma", "")).strip().upper()
        row_proj = str(row.get("Projeto", "")).strip().upper()

        pertence_aqui = (nome in nomes_atuais) or (row_plat == proj_norm) or (row_proj == proj_norm)
        if not pertence_aqui:
            continue

        if nome not in nomes_atuais:
            membros_atualizados.append({
                "nome": nome,
                "disciplina": disc,
                "funcao": str(row.get("Função", "Externo")).strip() if pd.notna(row.get("Função")) else "Externo",
                "projeto": row_proj if row_proj and row_proj != "NAN" else "EXTERNO",
                "ativo": 1,
                "is_externo": True
            })
            nomes_atuais.add(nome)

        # Mapeamento estrito solicitado pelo usuário
        if tipo_raw == "Estaleiro":
            st_crono = "BASE"
            obs_crono = "Estaleiro (Cronograma)"
        elif tipo_raw in ("Férias", "Ferias"):
            st_crono = "FÉR"
            obs_crono = "Férias (Cronograma)"
        else: # Folga
            st_crono = "FOL"
            obs_crono = "Folga (Cronograma)"

        d_ini_ev = max(row["_ini"].date(), dt_ini_mes)
        d_fim_ev = min(row["_fim"].date(), dt_fim_mes)

        curr = d_ini_ev
        while curr <= d_fim_ev:
            iso_d = curr.strftime("%Y-%m-%d")
            st_sqlite = merged_map.get((nome, iso_d))

            if st_sqlite is None:
                # Sem evento prévio na plataforma: plota o evento do cronograma diretamente
                merged_map[(nome, iso_d)] = st_crono
                merged_periodos[(nome, iso_d)] = {
                    "nome": nome,
                    "disciplina": disc,
                    "projeto": proj_norm,
                    "inicio": row["_ini"].date(),
                    "fim": row["_fim"].date(),
                    "status": st_crono,
                    "dias": (row["_fim"].date() - row["_ini"].date()).days + 1,
                    "observacao": obs_crono,
                    "origem_cronograma": True
                }
            elif st_sqlite != st_crono and st_sqlite != "CONF":
                # Conflito detectado entre escala e cronograma!
                merged_map[(nome, iso_d)] = "CONF"
                desc_conf = f"Escala {proj_norm} ({st_sqlite}) vs Cronograma ({tipo_raw})"

                # Preserva o período original da escala em SQLite para manter os limites reais (início e término)
                periodo_base = mapa_dia_periodo.get((nome, iso_d))
                if periodo_base and "inicio" in periodo_base and "fim" in periodo_base:
                    p_ini_real = periodo_base["inicio"]
                    p_fim_real = periodo_base["fim"]
                    p_st_real = periodo_base.get("status_escala") or periodo_base.get("status") or st_sqlite
                    p_dias_real = periodo_base.get("dias") or ((p_fim_real - p_ini_real).days + 1)
                else:
                    p_ini_real = curr
                    p_fim_real = curr
                    p_st_real = st_sqlite
                    p_dias_real = 1

                merged_periodos[(nome, iso_d)] = {
                    "nome": nome,
                    "disciplina": disc,
                    "projeto": proj_norm,
                    "inicio": p_ini_real,
                    "fim": p_fim_real,
                    "status": p_st_real,
                    "status_escala": p_st_real,
                    "dias": p_dias_real,
                    "observacao": desc_conf,
                    "is_conflito": True,
                    "status_crono": st_crono,
                    "tipo_crono": tipo_raw,
                    "origem_cronograma": False
                }
                lista_conflitos.append({
                    "nome": nome,
                    "data": iso_d,
                    "status_escala": st_sqlite,
                    "status_crono": st_crono,
                    "tipo_crono": tipo_raw,
                    "detalhe": f"CONFLITO em {iso_d}: {desc_conf}"
                })

            curr += timedelta(days=1)

    ordem_disc = {d: i for i, d in enumerate(DISCIPLINAS_PADRAO)}
    membros_atualizados.sort(
        key=lambda x: (ordem_disc.get(x["disciplina"], 99), x["nome"])
    )
    qtd_ext = sum(1 for m in membros_atualizados if m.get("is_externo"))

    return merged_map, merged_periodos, lista_conflitos, membros_atualizados, qtd_ext


# --- GERADOR DE HTML DA GRADE DE ESCALA (JANELA CONTÍNUA 6 MESES / RESUMO VAGAS / STATUS LIC & EMP) ---
def build_scala_html(
    projeto: str,
    ano: int,
    mes: int,
    membros: List[Dict],
    escalas_map: Dict[Tuple[str, str], str],
    mapa_dia_periodo: Dict[Tuple[str, str], Dict],
    qtd_nativos: int,
    qtd_externos: int,
    is_editor: bool,
    filtro_busca: str = "",
    meses_info: Optional[List[Dict]] = None,
    lista_dias: Optional[List[date]] = None,
    config_vagas: Optional[Dict[str, int]] = None
) -> str:
    """
    Gera o HTML compacto e fluido da escala:
    - Janela contínua de 6 meses com rolagem horizontal contínua
    - Cabeçalho duplo fixo (Sticky): Linha 1 = Mês/Ano; Linha 2 = Dias do mês com indicador do dia da semana
    - Coluna de nomes e disciplinas fixa à esquerda (Sticky)
    - Delimitação visual nítida no último dia de cada mês
    - Suporte aos novos status: LIC (Licença, #9f1239) e EMP (Emprestado, #4338ca)
    - Seleção de Regime de Folga (14x21 vs 14x14)
    - Resumo diário de efetivo a bordo com formatação condicional de alerta para vagas mínimas (Mecânica, Instrumentação, Elétrica e Total)
    - Eletricistas identificados em Amarelo (#eab308) na divisão de turma a bordo
    - Modais interativos para edição, exclusão, auto-permuta e pop-up informativo de colaborador emprestado (EMP)
    """
    hoje_iso = date.today().strftime("%Y-%m-%d")

    if not lista_dias:
        _, num_dias = calendar.monthrange(ano, mes)
        lista_dias = [date(ano, mes, d) for d in range(1, num_dias + 1)]
        meses_info = [{
            "ano": ano,
            "mes": mes,
            "nome": MESES_PT[mes - 1],
            "dias": num_dias
        }]
    elif not meses_info:
        meses_dict = {}
        for d in lista_dias:
            k = (d.year, d.month)
            if k not in meses_dict:
                meses_dict[k] = {"ano": d.year, "mes": d.month, "nome": MESES_PT[d.month - 1], "dias": 0}
            meses_dict[k]["dias"] += 1
        meses_info = list(meses_dict.values())

    config_vagas = config_vagas or {"min_mecanica": 0, "min_instrumentacao": 0, "min_eletrica": 0, "min_total": 0}
    min_mec = int(config_vagas.get("min_mecanica", 0) or 0)
    min_inst = int(config_vagas.get("min_instrumentacao", 0) or 0)
    min_elet = int(config_vagas.get("min_eletrica", 0) or 0)
    min_tot = int(config_vagas.get("min_total", 0) or 0)

    filtro = filtro_busca.strip().upper()
    if filtro:
        membros_filtrados = [
            m for m in membros
            if filtro in m["nome"] or filtro in m["disciplina"].upper()
        ]
    else:
        membros_filtrados = membros

    grupos: Dict[str, List[Dict]] = {d: [] for d in DISCIPLINAS_PADRAO}
    for m in membros_filtrados:
        disc = m["disciplina"]
        if is_mecanica(disc):
            grupos["Mecânica"].append(m)
        elif is_instrumentacao(disc):
            grupos["Instrumentação"].append(m)
        elif is_eletrica(disc):
            grupos["Elétrica"].append(m)
        else:
            if disc not in grupos:
                grupos[disc] = []
            grupos[disc].append(m)

    total_dias = len(lista_dias)
    col_width_dia = 37
    col_width_nome = 200
    total_grid_min_width = col_width_nome + (total_dias * col_width_dia)

    css = f"""
    <style>
        * {{ box-sizing: border-box; }}
        body {{
            margin: 0;
            padding: 8px;
            font-family: -apple-system, BlinkMacSystemFont, Arial, sans-serif;
            background: #f3f6f8;
            color: #24323d;
        }}
        .header-bar {{
            background: #075b9b;
            color: #ffffff;
            padding: 10px 14px;
            font-size: 15px;
            font-weight: bold;
            border-radius: 6px 6px 0 0;
            display: flex;
            justify-content: space-between;
            align-items: center;
        }}
        .legenda-bar {{
            display: flex;
            flex-wrap: wrap;
            gap: 10px;
            align-items: center;
            padding: 6px 12px;
            background: #ffffff;
            border-bottom: 1px solid #d7e0e6;
            font-size: 11px;
        }}
        .legenda-tag {{
            display: inline-flex;
            align-items: center;
            gap: 4px;
        }}
        .legenda-color {{
            display: inline-block;
            width: 12px;
            height: 12px;
            border-radius: 2px;
        }}
        .calendar {{
            overflow: auto;
            background: #ffffff;
            border: 1px solid #d7e0e6;
            border-radius: 0 0 6px 6px;
            max-height: 740px;
        }}
        .grid {{
            display: grid;
            grid-template-columns: {col_width_nome}px repeat({total_dias}, {col_width_dia}px);
            min-width: {total_grid_min_width}px;
        }}
        .cell {{
            min-height: 34px;
            padding: 2px 1px;
            border-right: 1px solid #dfe6eb;
            border-bottom: 1px solid #dfe6eb;
            font-size: 11px;
            display: flex;
            flex-direction: column;
            justify-content: center;
            align-items: center;
            position: relative;
        }}
        .head {{
            background: #eaf2f8;
            font-weight: bold;
            text-align: center;
            position: sticky;
            padding: 2px 1px;
        }}
        .label {{
            position: sticky;
            left: 0;
            background: #ffffff;
            z-index: 2;
            font-weight: bold;
            justify-content: center;
            align-items: flex-start;
            padding-left: 8px;
            border-right: 2px solid #cbd5e1;
            white-space: nowrap;
            overflow: hidden;
            text-overflow: ellipsis;
        }}
        .head.label {{
            background: #eaf2f8;
            align-items: center;
            justify-content: center;
            padding-left: 0;
        }}
        .weekend {{
            background: #f7f7f7;
        }}
        .today {{
            background: #fef2f2 !important;
            border-left: 2px solid #dc2626 !important;
            border-right: 2px solid #dc2626 !important;
        }}
        .border-month-end {{
            border-right: 2px solid #0284c7 !important;
        }}
        .discipline-header {{
            grid-column: 1 / -1;
            background: #075b9b;
            color: #ffffff;
            font-weight: bold;
            padding: 6px 12px;
            font-size: 12px;
            text-transform: uppercase;
            letter-spacing: 0.5px;
            position: sticky;
            left: 0;
            z-index: 2;
        }}
        .item-compact {{
            border-radius: 3px;
            padding: 3px 1px;
            width: 94%;
            text-align: center;
            box-shadow: 0 1px 2px rgba(0,0,0,0.12);
            font-weight: bold;
            font-size: 10px;
            letter-spacing: 0px;
            white-space: nowrap;
            overflow: hidden;
            text-overflow: ellipsis;
            line-height: 1.1;
            display: block;
            text-decoration: none !important;
        }}
        .item-clickable {{
            cursor: pointer;
            transition: transform 0.1s, box-shadow 0.1s;
        }}
        .item-clickable:hover {{
            transform: scale(1.15);
            box-shadow: 0 2px 5px rgba(0,0,0,0.3);
            outline: 2px solid #ffffff;
            z-index: 5;
        }}
        .item-conf {{
            background: #dc2626 !important;
            color: #ffffff !important;
            font-weight: 800 !important;
            border: 1px solid #991b1b !important;
            box-shadow: 0 1px 3px rgba(220, 38, 38, 0.45);
            cursor: pointer;
        }}
        .item-conf:hover {{
            background: #b91c1c !important;
            outline: 2px solid #fecaca !important;
            transform: scale(1.18);
            z-index: 10;
        }}
        .item-emp {{
            background: #4338ca !important;
            color: #ffffff !important;
            border: 1px solid #312e81 !important;
            cursor: pointer;
        }}
        .item-emp:hover {{
            background: #3730a3 !important;
            outline: 2px solid #c7d2fe !important;
            transform: scale(1.18);
            z-index: 10;
        }}
        .empty-dash {{
            color: #e2e8f0;
            font-size: 9px;
        }}
        .empty-clickable {{
            color: #cbd5e1;
            font-size: 11px;
            width: 100%;
            height: 100%;
            display: flex;
            align-items: center;
            justify-content: center;
            cursor: pointer;
            border-radius: 2px;
            transition: background 0.12s, color 0.12s;
        }}
        .empty-clickable:hover {{
            background: #e0f2fe;
            color: #0284c7;
            font-weight: bold;
        }}
        .badge-externo {{
            background: #fef3c7;
            color: #92400e;
            border: 1px solid #f59e0b;
            font-size: 9px;
            padding: 1px 3px;
            border-radius: 3px;
            margin-top: 1px;
            font-weight: 600;
        }}
        .hint-edicao {{
            background: #fef9c3;
            border-bottom: 1px solid #fde047;
            padding: 6px 14px;
            font-size: 11px;
            color: #854d0e;
            display: flex;
            align-items: center;
            gap: 6px;
        }}
        .turma-header {{
            background: #0f172a !important;
            color: #f8fafc !important;
            font-size: 12px !important;
            display: flex;
            justify-content: space-between;
            align-items: center;
            border-top: 3px solid #0284c7;
            margin-top: 10px;
        }}
        .turma-legenda-mini {{
            display: flex;
            gap: 6px;
            font-size: 10px;
            font-weight: normal;
        }}
        .turma-chip-legenda {{
            padding: 2px 7px;
            border-radius: 4px;
            color: #ffffff;
            font-weight: 600;
        }}
        .turma-balao {{
            display: flex;
            align-items: center;
            justify-content: center;
            margin: 2px 1px;
            padding: 0 6px;
            border-radius: 12px;
            font-size: 10px;
            font-weight: 700;
            white-space: nowrap;
            overflow: hidden;
            text-overflow: ellipsis;
            box-shadow: 0 1px 3px rgba(0,0,0,0.15);
            z-index: 1;
            min-height: 26px;
            box-sizing: border-box;
            border-bottom: 1px solid #dfe6eb;
            transition: filter 0.15s ease, transform 0.15s ease;
        }}
        .turma-balao:hover {{
            filter: brightness(1.08);
            transform: translateY(-1px);
        }}
        .vagas-alerta {{
            background: #fee2e2 !important;
            color: #b91c1c !important;
            border: 1px solid #fca5a5 !important;
            font-weight: 800 !important;
            font-size: 11.5px !important;
        }}
        /* Modal Styles */
        .p80-modal-overlay {{
            position: fixed;
            top: 0;
            left: 0;
            width: 100vw;
            height: 100vh;
            background: rgba(15, 23, 42, 0.65);
            backdrop-filter: blur(4px);
            -webkit-backdrop-filter: blur(4px);
            display: none;
            align-items: center;
            justify-content: center;
            z-index: 99999;
        }}
        .p80-modal-card {{
            background: #ffffff;
            border-radius: 12px;
            width: 90%;
            max-width: 490px;
            box-shadow: 0 25px 50px -12px rgba(0, 0, 0, 0.4);
            overflow: hidden;
            display: flex;
            flex-direction: column;
            animation: p80PopIn 0.18s cubic-bezier(0.16, 1, 0.3, 1);
        }}
        @keyframes p80PopIn {{
            from {{ opacity: 0; transform: scale(0.92) translateY(8px); }}
            to {{ opacity: 1; transform: scale(1) translateY(0); }}
        }}
        .p80-modal-head {{
            background: #075b9b;
            color: #ffffff;
            padding: 14px 18px;
            display: flex;
            justify-content: space-between;
            align-items: center;
        }}
        .p80-modal-head h3 {{
            margin: 0;
            font-size: 15px;
            font-weight: 700;
        }}
        .p80-modal-subtitle {{
            font-size: 11px;
            opacity: 0.9;
            margin-top: 3px;
            font-weight: normal;
        }}
        .p80-modal-close-btn {{
            background: transparent;
            border: none;
            color: #ffffff;
            font-size: 26px;
            line-height: 1;
            cursor: pointer;
            opacity: 0.85;
            padding: 0 4px;
            transition: opacity 0.15s, transform 0.15s;
        }}
        .p80-modal-close-btn:hover {{
            opacity: 1;
            transform: scale(1.15);
        }}
        .p80-modal-body {{
            padding: 16px 20px;
            display: flex;
            flex-direction: column;
            gap: 12px;
            background: #ffffff;
        }}
        .p80-form-group {{
            display: flex;
            flex-direction: column;
            gap: 4px;
        }}
        .p80-form-group label {{
            font-size: 12px;
            font-weight: 600;
            color: #334155;
        }}
        .p80-form-row {{
            display: grid;
            grid-template-columns: 1fr 1fr;
            gap: 12px;
        }}
        .p80-input, .p80-select {{
            padding: 8px 10px;
            border: 1px solid #cbd5e1;
            border-radius: 6px;
            font-size: 13px;
            color: #1e293b;
            background: #f8fafc;
            outline: none;
            transition: border-color 0.15s, box-shadow 0.15s;
            font-family: inherit;
            box-sizing: border-box;
            width: 100%;
        }}
        .p80-input:focus, .p80-select:focus {{
            border-color: #075b9b;
            box-shadow: 0 0 0 3px rgba(7, 91, 155, 0.15);
            background: #ffffff;
        }}
        .p80-info-box {{
            background: #eff6ff;
            border: 1px solid #bfdbfe;
            border-radius: 6px;
            padding: 10px 12px;
            font-size: 11.5px;
            color: #1e40af;
            line-height: 1.4;
        }}
        .p80-modal-foot {{
            background: #f8fafc;
            border-top: 1px solid #e2e8f0;
            padding: 12px 20px;
            display: flex;
            justify-content: flex-end;
            gap: 8px;
            align-items: center;
        }}
        .p80-btn {{
            padding: 8px 14px;
            border-radius: 6px;
            font-size: 12.5px;
            font-weight: 600;
            cursor: pointer;
            border: none;
            transition: all 0.15s ease;
            display: inline-flex;
            align-items: center;
            gap: 5px;
            text-decoration: none;
        }}
        .p80-btn-primary {{
            background: #075b9b;
            color: #ffffff;
        }}
        .p80-btn-primary:hover {{
            background: #05487a;
            box-shadow: 0 2px 4px rgba(7, 91, 155, 0.25);
        }}
        .p80-btn-danger {{
            background: #fee2e2;
            color: #b91c1c;
            border: 1px solid #fca5a5;
        }}
        .p80-btn-danger:hover {{
            background: #fecaca;
        }}
        .p80-btn-cancel {{
            background: #e2e8f0;
            color: #475569;
        }}
        .p80-btn-cancel:hover {{
            background: #cbd5e1;
        }}
    </style>
    """

    legenda_html = '<div class="legenda-bar"><strong>Legenda:</strong>'
    for k, v in STATUS_MAP.items():
        if k in ("FÉRIAS", "TREINAMENTO", "E", "CONFLITO", "ESTALEIRO", "LICENÇA", "LICENCA", "EMPRESTADO"):
            continue
        legenda_texto = v['codigo'] if v['codigo'] == v['nome'] else f"{v['codigo']} ({v['nome']})"
        legenda_html += f"""
        <div class="legenda-tag">
            <span class="legenda-color" style="background:{v['bg']};"></span>
            <span><strong>{legenda_texto}</strong></span>
        </div>
        """
    legenda_html += '</div>'

    hint_html = ""
    if is_editor:
        hint_html = '<div class="hint-edicao">💡 <strong>Modo Edição Ativo:</strong> Clique em qualquer período colorido para editar datas, regime ou condição, ou clique em um dia livre (—) para lançar escala.</div>'

    grid_cells = []

    # Linha 1: Mês / Ano (Banner superior agrupando os dias de cada mês)
    grid_cells.append(
        '<div class="cell head label" style="grid-column: 1 / 2; top: 0; z-index: 6; background: #0369a1; color: #ffffff; font-size: 11px; text-transform: uppercase; border-bottom: 1px solid #0284c7; height: 30px; min-height: 30px;">MÊS / ANO</div>'
    )
    for m_inf in meses_info:
        m_span = m_inf["dias"]
        grid_cells.append(
            f'<div class="cell head" style="grid-column: span {m_span}; top: 0; z-index: 5; background: #075b9b; color: #ffffff; font-weight: bold; font-size: 11.5px; letter-spacing: 0.5px; border-right: 2px solid #0284c7; border-bottom: 1px solid #0284c7; height: 30px; min-height: 30px; display: flex; align-items: center; justify-content: center;">'
            f'📅 {m_inf["nome"].upper()} {m_inf["ano"]}'
            f'</div>'
        )

    # Linha 2: Dias do Mês e Dias da Semana
    grid_cells.append(
        '<div class="cell head label" style="top: 30px; z-index: 5; height: 36px; min-height: 36px;">PROFISSIONAL / EQUIPE</div>'
    )
    for idx_dia, dt in enumerate(lista_dias):
        sem_idx = dt.weekday()
        sem_str = DIAS_SEMANA_PT[sem_idx]
        is_wk = sem_idx in [5, 6]
        wk_cls = " weekend" if is_wk else ""
        today_cls = " today" if dt.strftime("%Y-%m-%d") == hoje_iso else ""
        is_month_end = (idx_dia == total_dias - 1) or (lista_dias[idx_dia + 1].month != dt.month)
        month_end_cls = " border-month-end" if is_month_end else ""

        grid_cells.append(
            f'<div class="cell head{wk_cls}{today_cls}{month_end_cls}" style="top: 30px; z-index: 4; height: 36px; min-height: 36px;">'
            f'{dt.day:02d}<br><small style="font-size:8px;color:#64748b;">{sem_str}</small>'
            f'</div>'
        )

    # Linhas dos Colaboradores
    total_linhas = 0
    for disc, lista_membros in grupos.items():
        if not lista_membros:
            continue

        grid_cells.append(
            f'<div class="discipline-header">📁 {disc.upper()} ({len(lista_membros)} colaboradores)</div>'
        )

        for m in lista_membros:
            total_linhas += 1
            nome = m["nome"]
            is_ext = m.get("is_externo", False)
            reg_colab = m.get("regime_folga", "14x21")
            reg_badge = f'<span style="font-size:8.5px;background:#e0f2fe;color:#0369a1;padding:1px 4px;border-radius:3px;margin-left:4px;font-weight:600;">{reg_colab}</span>'
            badge_origem = f'<span class="badge-externo">Externo ({m.get("projeto", "Outro")})</span>' if is_ext else ""

            grid_cells.append(
                f'<div class="cell label"><span>{nome}</span><small style="color:#64748b;font-weight:normal;font-size:9px;">{disc}{reg_badge}</small>{badge_origem}</div>'
            )

            nome_js = nome.replace("'", "\\'").replace('"', '&quot;')
            disc_js = disc.replace("'", "\\'").replace('"', '&quot;')

            for idx_dia, dt in enumerate(lista_dias):
                iso_dt = dt.strftime("%Y-%m-%d")
                is_wk = dt.weekday() in [5, 6]
                wk_cls = " weekend" if is_wk else ""
                today_cls = " today" if iso_dt == hoje_iso else ""
                is_month_end = (idx_dia == total_dias - 1) or (lista_dias[idx_dia + 1].month != dt.month)
                month_end_cls = " border-month-end" if is_month_end else ""

                status_raw = escalas_map.get((nome, iso_dt))
                status = normalizar_status(status_raw) if status_raw else None

                if status and status in STATUS_MAP:
                    st_info = STATUS_MAP[status]
                    periodo_obj = mapa_dia_periodo.get((nome, iso_dt))
                    obs_val = (periodo_obj.get("observacao") if periodo_obj else "") or ""

                    if status == "EMP":
                        dest_proj = (periodo_obj.get("projeto_destino") if periodo_obj else None) or "Outra Unidade"
                        p_ini_f = periodo_obj["inicio"].strftime("%d/%m/%Y") if periodo_obj else dt.strftime("%d/%m/%Y")
                        p_fim_f = periodo_obj["fim"].strftime("%d/%m/%Y") if periodo_obj else dt.strftime("%d/%m/%Y")
                        tooltip = f"Emprestado para {dest_proj} ({p_ini_f} a {p_fim_f}) — Clique para ver detalhes"
                        tooltip_esc = tooltip.replace('"', '&quot;')
                        dest_proj_js = dest_proj.replace("'", "\\'").replace('"', '&quot;')
                        cell_content = f"""
                        <div class="item-compact item-clickable item-emp"
                             onclick="abrirModalEmprestado('{nome_js}', '{dest_proj_js}', '{p_ini_f}', '{p_fim_f}')"
                             style="background:{st_info['bg']};color:{st_info['color']};"
                             title="{tooltip_esc}">
                            {st_info['codigo']}
                        </div>
                        """
                    elif status == "CONF":
                        desc_conf = obs_val if obs_val else f"Divergência entre Escala {projeto} e Cronograma"
                        tooltip = f"⚠️ CONFLITO em {dt.day:02d}/{dt.month:02d}: {desc_conf}"
                        if is_editor and periodo_obj:
                            p_st_real = periodo_obj.get("status_escala") or periodo_obj.get("status") or "EMB"
                            tooltip += f" | Escala: {periodo_obj['inicio']:%d/%m} a {periodo_obj['fim']:%d/%m} ({p_st_real}) (Clique para editar/excluir)"
                            p_ini_str = periodo_obj["inicio"].strftime("%Y-%m-%d")
                            p_fim_str = periodo_obj["fim"].strftime("%Y-%m-%d")
                            obs_js = obs_val.replace("'", "\'")
                            tooltip_esc = tooltip.replace('"', '&quot;')
                            cell_content = f"""
                            <div class="item-compact item-clickable item-conf"
                                 onclick="abrirModalEdicao('{nome_js}', '{disc_js}', '{projeto}', '{p_ini_str}', '{p_fim_str}', '{p_st_real}', '{obs_js}', false, '{reg_colab}')"
                                 title="{tooltip_esc}">
                                CONF
                            </div>
                            """
                        else:
                            tooltip_esc = tooltip.replace('"', '&quot;')
                            cell_content = f'<div class="item-compact item-conf" title="{tooltip_esc}">CONF</div>'
                    else:
                        if is_editor and periodo_obj:
                            p_st = periodo_obj.get("status_escala") or periodo_obj.get("status") or status
                            p_ini_str = periodo_obj["inicio"].strftime("%Y-%m-%d")
                            p_fim_str = periodo_obj["fim"].strftime("%Y-%m-%d")
                            tooltip = f"Clique para editar: {nome} | {periodo_obj['inicio']:%d/%m} a {periodo_obj['fim']:%d/%m} ({p_st})"
                            if obs_val:
                                tooltip += f" — {obs_val}"
                            tooltip_esc = tooltip.replace('"', '&quot;')
                            obs_js = obs_val.replace("'", "\'")
                            cell_content = f"""
                            <div class="item-compact item-clickable"
                                 onclick="abrirModalEdicao('{nome_js}', '{disc_js}', '{projeto}', '{p_ini_str}', '{p_fim_str}', '{p_st}', '{obs_js}', false, '{reg_colab}')"
                                 style="background:{st_info['bg']};color:{st_info['color']};"
                                 title="{tooltip_esc}">
                                {st_info['codigo']}
                            </div>
                            """
                        else:
                            tooltip = f"{nome} em {dt.day:02d}/{dt.month:02d}: {st_info['nome']}"
                            if obs_val:
                                tooltip += f" ({obs_val})"
                            tooltip_esc = tooltip.replace('"', '&quot;')
                            cell_content = f"""
                            <div class="item-compact" style="background:{st_info['bg']};color:{st_info['color']};" title="{tooltip_esc}">
                                {st_info['codigo']}
                            </div>
                            """
                else:
                    if is_editor:
                        dt_fim_sugest = (dt + timedelta(days=13)).strftime("%Y-%m-%d")
                        tooltip_vazio = f"Clique para lançar escala para {nome} a partir de {dt.day:02d}/{dt.month:02d}"
                        cell_content = f"""
                        <div class="empty-clickable"
                             onclick="abrirModalEdicao('{nome_js}', '{disc_js}', '{projeto}', '{iso_dt}', '{dt_fim_sugest}', 'EMB', '', true, '{reg_colab}')"
                             title="{tooltip_vazio}">
                            —
                        </div>
                        """
                    else:
                        cell_content = '<span class="empty-dash">—</span>'

                grid_cells.append(f'<div class="cell{wk_cls}{today_cls}{month_end_cls}">{cell_content}</div>')

    # --- DIVISÃO DE TURMA A BORDO (AO FINAL DE TODAS AS DISCIPLINAS) ---
    def _is_embarcado_dia(p_nome: str, p_iso: str) -> bool:
        st_val = escalas_map.get((p_nome, p_iso))
        if st_val == "EMP":
            return False  # Emprestado para outra plataforma; NÃO conta nesta unidade!
        if st_val in ("EMB", "E"):
            return True
        if st_val == "CONF":
            p_obj = mapa_dia_periodo.get((p_nome, p_iso))
            if p_obj and (p_obj.get("status_escala") in ("EMB", "E") or p_obj.get("status") in ("EMB", "E")):
                return True
        return False

    soma_total_dia = {}
    soma_mec_dia = {}
    soma_inst_dia = {}
    soma_elet_dia = {}

    for idx_dia, dt in enumerate(lista_dias):
        iso_dt = dt.strftime("%Y-%m-%d")
        tot = 0
        mec = 0
        inst = 0
        elet = 0
        for m in membros_filtrados:
            if _is_embarcado_dia(m["nome"], iso_dt):
                tot += 1
                d_m = m.get("disciplina", "")
                if is_mecanica(d_m):
                    mec += 1
                elif is_instrumentacao(d_m):
                    inst += 1
                elif is_eletrica(d_m):
                    elet += 1
        soma_total_dia[idx_dia] = tot
        soma_mec_dia[idx_dia] = mec
        soma_inst_dia[idx_dia] = inst
        soma_elet_dia[idx_dia] = elet

    embarques_turma = []
    for m in membros_filtrados:
        nome = m["nome"]
        disc = m["disciplina"]
        cols_emb = [idx for idx, dt in enumerate(lista_dias) if _is_embarcado_dia(nome, dt.strftime("%Y-%m-%d"))]
        if not cols_emb:
            continue
        blocos = []
        b_ini = cols_emb[0]
        b_fim = cols_emb[0]
        for c in cols_emb[1:]:
            if c == b_fim + 1:
                b_fim = c
            else:
                blocos.append((b_ini, b_fim))
                b_ini = c
                b_fim = c
        blocos.append((b_ini, b_fim))
        embarques_turma.append({
            "nome": nome,
            "disciplina": disc,
            "is_externo": m.get("is_externo", False),
            "projeto": m.get("projeto", "Outro"),
            "regime_folga": m.get("regime_folga", "14x21"),
            "blocos": blocos
        })

    def _ord_disc(d):
        if is_mecanica(d): return 1
        if is_instrumentacao(d): return 2
        if is_eletrica(d): return 3
        return 4

    embarques_turma.sort(key=lambda x: (_ord_disc(x["disciplina"]), x["blocos"][0][0], x["nome"]))

    # 1. Cabeçalho da Divisão de Turma a Bordo
    grid_cells.append(
        f'<div class="discipline-header turma-header">'
        f'<span>🚢 DIVISÃO DE TURMA A BORDO ({len(embarques_turma)} profissionais a bordo no período)</span>'
        f'<span class="turma-legenda-mini">'
        f'<span class="turma-chip-legenda" style="background:#0284c7;">Mecânica</span>'
        f'<span class="turma-chip-legenda" style="background:#ea580c;">Instrumentação</span>'
        f'<span class="turma-chip-legenda" style="background:#eab308;color:#1e293b;">Elétrica</span>'
        f'</span>'
        f'</div>'
    )

    # 2. Linhas de Resumo Diário com Formatação Condicional de Metas Mínimas
    def _render_linha_resumo(titulo: str, dict_soma: Dict[int, int], min_val: int, cor_bg: str, cor_txt: str):
        meta_badge = f' <span style="font-size:9px;background:#ffffff;padding:1px 5px;border-radius:3px;border:1px solid #cbd5e1;color:#334155;">Mín: {min_val}</span>' if min_val > 0 else ""
        grid_cells.append(
            f'<div class="cell head label" style="background:{cor_bg};color:{cor_txt};font-weight:700;font-size:10.5px;border-bottom:1px solid #cbd5e1;">'
            f'{titulo}{meta_badge}</div>'
        )
        for idx_dia, dt in enumerate(lista_dias):
            qtd = dict_soma[idx_dia]
            iso_dt = dt.strftime("%Y-%m-%d")
            is_wk = dt.weekday() in [5, 6]
            wk_cls = " weekend" if is_wk else ""
            today_cls = " today" if iso_dt == hoje_iso else ""
            is_month_end = (idx_dia == total_dias - 1) or (lista_dias[idx_dia + 1].month != dt.month)
            month_end_cls = " border-month-end" if is_month_end else ""

            if min_val > 0 and qtd < min_val:
                grid_cells.append(
                    f'<div class="cell head vagas-alerta{wk_cls}{today_cls}{month_end_cls}" '
                    f'title="⚠️ Alerta: {qtd} a bordo em {dt.strftime("%d/%m/%Y")} (Mínimo exigido: {min_val})">'
                    f'⚠️{qtd}</div>'
                )
            elif qtd > 0:
                grid_cells.append(
                    f'<div class="cell head{wk_cls}{today_cls}{month_end_cls}" style="background:#f0f9ff;color:{cor_txt};font-weight:bold;font-size:12px;border-bottom:1px solid #cbd5e1;" '
                    f'title="{qtd} em {dt.strftime("%d/%m/%Y")}">{qtd}</div>'
                )
            else:
                grid_cells.append(
                    f'<div class="cell head{wk_cls}{today_cls}{month_end_cls}" style="background:#f8fafc;color:#94a3b8;font-size:11px;border-bottom:1px solid #cbd5e1;">0</div>'
                )

    _render_linha_resumo("📊 TOTAL A BORDO (SOMA)", soma_total_dia, min_tot, "#e0f2fe", "#0369a1")
    _render_linha_resumo("🔧 MECÂNICA A BORDO", soma_mec_dia, min_mec, "#f0fdf4", "#15803d")
    _render_linha_resumo("⚡ ELÉTRICA A BORDO", soma_elet_dia, min_elet, "#fefce8", "#a16207")
    _render_linha_resumo("🎛️ INSTRUMENTAÇÃO A BORDO", soma_inst_dia, min_inst, "#fff7ed", "#c2410c")

    # 3. Linhas individuais com balão contínuo por período de embarque
    for item in embarques_turma:
        nome = item["nome"]
        disc = item["disciplina"]
        reg_colab = item.get("regime_folga", "14x21")

        if is_mecanica(disc):
            bg_balao = "#0284c7"
            borda_balao = "#0369a1"
            text_balao = "#ffffff"
        elif is_instrumentacao(disc):
            bg_balao = "#ea580c"
            borda_balao = "#c2410c"
            text_balao = "#ffffff"
        elif is_eletrica(disc):
            bg_balao = "#eab308"
            borda_balao = "#ca8a04"
            text_balao = "#1e293b"
        else:
            bg_balao = "#0d9488"
            borda_balao = "#0f766e"
            text_balao = "#ffffff"

        badge_ext = '<span class="badge-externo">Ext</span>' if item.get("is_externo") else ""
        grid_cells.append(
            f'<div class="cell label"><span>{nome}</span><small style="color:#64748b;font-weight:normal;font-size:9px;">{disc}</small>{badge_ext}</div>'
        )

        nome_js = nome.replace("'", "\'")
        disc_js = disc.replace("'", "\'")

        curr_col = 2
        for (b_ini, b_fim) in item["blocos"]:
            span_vazio = (b_ini + 2) - curr_col
            if span_vazio > 0:
                grid_cells.append(f'<div class="cell" style="grid-column: {curr_col} / span {span_vazio}; border-right:1px solid #dfe6eb; border-bottom:1px solid #dfe6eb;"></div>')

            span_balao = b_fim - b_ini + 1
            dt_b_ini = lista_dias[b_ini]
            dt_b_fim = lista_dias[b_fim]
            iso_b_ini = dt_b_ini.strftime("%Y-%m-%d")
            iso_b_fim = dt_b_fim.strftime("%Y-%m-%d")

            periodo_b = mapa_dia_periodo.get((nome, iso_b_ini))
            if periodo_b:
                p_ini_modal = periodo_b["inicio"].strftime("%Y-%m-%d")
                p_fim_modal = periodo_b["fim"].strftime("%Y-%m-%d")
                obs_modal = (periodo_b.get("observacao") or "").replace("'", "\'")
            else:
                p_ini_modal = iso_b_ini
                p_fim_modal = iso_b_fim
                obs_modal = ""

            tooltip_balao = f"{nome} ({disc}) — Embarcado de {dt_b_ini:%d/%m} a {dt_b_fim:%d/%m} ({span_balao} dias)"
            onclick_balao = f'onclick="abrirModalEdicao(\'{nome_js}\', \'{disc_js}\', \'{projeto}\', \'{p_ini_modal}\', \'{p_fim_modal}\', \'EMB\', \'{obs_modal}\', false, \'{reg_colab}\')"' if is_editor else ""
            cursor_balao = "cursor:pointer;" if is_editor else "cursor:default;"

            grid_cells.append(
                f'<div class="turma-balao" style="grid-column: {b_ini + 2} / span {span_balao}; background:{bg_balao}; border:1px solid {borda_balao}; color:{text_balao}; {cursor_balao}" '
                f'{onclick_balao} title="{tooltip_balao}">'
                f'<span>⚓ {nome} ({span_balao}d)</span>'
                f'</div>'
            )
            curr_col = b_fim + 3

        if curr_col <= total_dias + 1:
            span_fim = (total_dias + 2) - curr_col
            grid_cells.append(f'<div class="cell" style="grid-column: {curr_col} / span {span_fim}; border-right:1px solid #dfe6eb; border-bottom:1px solid #dfe6eb;"></div>')

    if not embarques_turma:
        grid_cells.append(
            f'<div class="cell label" style="grid-column: 1 / -1;text-align:center;padding:16px;color:#64748b;font-style:italic;">'
            f'Nenhum profissional com embarque ativo nesta unidade no período exibido.'
            f'</div>'
        )

    if total_linhas == 0:
        grid_cells.append(
            f'<div class="cell label" style="grid-column: 1 / -1;text-align:center;padding:25px;color:#64748b;">Nenhum profissional encontrado com os filtros atuais.</div>'
        )

    grid_body = "\n".join(grid_cells)

    periodo_rotulo = f"{meses_info[0]['nome']} de {meses_info[0]['ano']}" if len(meses_info) == 1 else f"{meses_info[0]['nome']}/{meses_info[0]['ano']} a {meses_info[-1]['nome']}/{meses_info[-1]['ano']} (6 Meses)"

    html_completo = f"""<!doctype html>
<html lang="pt-BR">
<head>
    <meta charset="utf-8">
    <meta name="viewport" content="width=device-width, initial-scale=1">
    <title>Escala {projeto}</title>
    {css}
</head>
<body>
    <div class="header-bar">
        <span>Escala {projeto} — {periodo_rotulo} <small style="font-size:11px;font-weight:normal;opacity:0.85;background:rgba(255,255,255,0.2);padding:2px 6px;border-radius:4px;margin-left:6px;">v0.2</small></span>
        <span style="font-size:12px;font-weight:normal;opacity:0.95;">
            Equipe Base: <strong>{qtd_nativos}</strong> | Externos: <strong>{qtd_externos}</strong> | Total Linhas: <strong>{total_linhas}</strong>
        </span>
    </div>
    {legenda_html}
    {hint_html}
    <div class="calendar">
        <div class="grid">
            {grid_body}
        </div>
    </div>

    <!-- MODAL POP-UP DE EDIÇÃO / AUTO-PERMUTA / LANÇAMENTO DE PERÍODO -->
    <div id="p80ModalOverlay" class="p80-modal-overlay">
        <div class="p80-modal-card">
            <div class="p80-modal-head" id="m_head_bar">
                <div>
                    <h3 id="m_title">✏️ Editar Período de Escala</h3>
                    <div class="p80-modal-subtitle" id="m_subtitle">Colaborador</div>
                </div>
                <button type="button" class="p80-modal-close-btn" onclick="fecharModal()" title="Fechar (Esc)">&times;</button>
            </div>

            <form id="modalForm" method="GET" target="_top">
                <input type="hidden" name="tab" id="m_tab_field" value="Escala {projeto}">
                <input type="hidden" name="p80_act" id="m_act" value="salvar">
                <input type="hidden" name="p80_proj" id="m_proj" value="{projeto}">
                <input type="hidden" name="p80_nome" id="m_nome" value="">
                <input type="hidden" name="p80_disc" id="m_disc" value="">
                <input type="hidden" name="p80_ini_ant" id="m_ini_ant" value="">
                <input type="hidden" name="p80_fim_ant" id="m_fim_ant" value="">
                <input type="hidden" name="p80_is_novo" id="m_is_novo" value="0">
                <input type="hidden" name="p80_mes" id="m_mes" value="{mes}">
                <input type="hidden" name="p80_ano" id="m_ano" value="{ano}">

                <!-- SEÇÃO 1: AUTO-PERMUTA DE FOLGA (Ativa automaticamente quando o status for FOL) -->
                <div id="secao_auto_permuta" class="p80-modal-body" style="display: none;">
                    <div style="background:#f0fdf4;border:1px solid #bbf7d0;padding:8px 12px;border-radius:6px;font-size:11.5px;color:#166534;display:flex;align-items:center;gap:6px;">
                        <span>🔄</span>
                        <span><strong>Janela de Auto-Permuta:</strong> Troque seu período de folga por outro de mesma duração contínua em até 180 dias.</span>
                    </div>

                    <div style="background:#f8fafc;border:1px solid #e2e8f0;padding:10px 12px;border-radius:6px;">
                        <span style="font-size:11px;font-weight:700;color:#475569;text-transform:uppercase;">1. Intervalo de Folga a Ceder (Origem):</span>
                        <div class="p80-form-row" style="margin-top:6px;">
                            <div class="p80-form-group">
                                <label for="ap_orig_ini">Data Início Origem:</label>
                                <input type="date" class="p80-input" name="ap_orig_ini" id="ap_orig_ini" onchange="atualizarAutoPermuta()">
                            </div>
                            <div class="p80-form-group">
                                <label for="ap_orig_fim">Data Fim Origem:</label>
                                <input type="date" class="p80-input" name="ap_orig_fim" id="ap_orig_fim" onchange="atualizarAutoPermuta()">
                            </div>
                        </div>
                        <div style="margin-top:4px;font-size:11.5px;color:#0369a1;font-weight:600;">
                            Folga a permutar: <span id="ap_dias_count">0</span> dia(s) contínuo(s).
                        </div>
                    </div>

                    <div style="background:#eff6ff;border:1px solid #bfdbfe;padding:10px 12px;border-radius:6px;">
                        <span style="font-size:11px;font-weight:700;color:#1e40af;text-transform:uppercase;">2. Nova Folga (Destino — Máx. 180 dias):</span>
                        <div class="p80-form-row" style="margin-top:6px;">
                            <div class="p80-form-group">
                                <label for="ap_dest_ini">Nova Data Início:</label>
                                <input type="date" class="p80-input" name="ap_dest_ini" id="ap_dest_ini" onchange="atualizarAutoPermuta()">
                            </div>
                            <div class="p80-form-group">
                                <label for="ap_dest_fim">Nova Data Término (Automática):</label>
                                <input type="date" class="p80-input" name="ap_dest_fim" id="ap_dest_fim" readonly style="background:#f1f5f9;cursor:not-allowed;">
                            </div>
                        </div>
                    </div>

                    <div id="ap_feedback_box" class="p80-info-box" style="display:none;"></div>
                </div>

                <!-- SEÇÃO 2: EDIÇÃO REGULAR DE PERÍODO (Ativa para EMB, PMB, FÉR, BASE, TREIN, LIC, CONF ou Novo) -->
                <div id="secao_edicao_padrao" class="p80-modal-body">
                    <div class="p80-form-row">
                        <div class="p80-form-group">
                            <label for="m_dt_ini">📅 Data Início:</label>
                            <input type="date" class="p80-input" name="p80_ini" id="m_dt_ini">
                        </div>
                        <div class="p80-form-group">
                            <label for="m_dt_fim">📅 Data Término:</label>
                            <input type="date" class="p80-input" name="p80_fim" id="m_dt_fim">
                        </div>
                    </div>

                    <div class="p80-form-group">
                        <label for="m_status">🏷️ Condição / Status:</label>
                        <select class="p80-select" name="p80_st" id="m_status">
                            <option value="EMB">EMB - Embarcado</option>
                            <option value="FOL">FOL - Folga</option>
                            <option value="PMB">PMB - PMB</option>
                            <option value="FÉR">FÉR - Férias</option>
                            <option value="BASE">BASE - Escritório / Base / Estaleiro</option>
                            <option value="TREIN">TREIN - Treinamento</option>
                            <option value="LIC">LIC - Licença</option>
                        </select>
                    </div>

                    <!-- SELEÇÃO DE REGIME DE FOLGA (14x21 vs 14x14) -->
                    <div class="p80-form-group" id="box_regime_folga">
                        <label for="m_regime_folga">⚖️ Projeção de Regime de Folga:</label>
                        <select class="p80-select" name="p80_regime" id="m_regime_folga" onchange="atualizarRegraFolga()">
                            <option value="14x21">14x21 (14 dias trab. / 21 dias folga — Fator 1.5x)</option>
                            <option value="14x14">14x14 (14 dias trab. / 14 dias folga — Fator 1.0x)</option>
                        </select>
                    </div>

                    <div class="p80-form-group">
                        <label for="m_obs">📝 Observação (opcional):</label>
                        <input type="text" class="p80-input" name="p80_obs" id="m_obs" placeholder="Ex: Embarque regular, estaleiro, licença médica...">
                    </div>

                    <div id="m_info_box" class="p80-info-box">
                        ⚓ <strong>Regra 1,5x:</strong> Embarques (EMB) geram automaticamente dias de folga (FOL) na sequência.
                    </div>

                    <!-- OPÇÃO DE ALOCAÇÃO EM ESCALA CÍCLICA -->
                    <div id="secao_opcao_ciclica" style="display:none;margin-top:8px;padding:10px 12px;background:#f0fdf4;border:1px solid #bbf7d0;border-radius:6px;">
                        <div style="display:flex;align-items:center;gap:8px;">
                            <input type="checkbox" id="m_chk_ciclica" name="p80_is_ciclica" value="1" onchange="atualizarOpcaoCiclica()" style="width:16px;height:16px;cursor:pointer;">
                            <label for="m_chk_ciclica" style="font-size:12px;font-weight:700;color:#166534;cursor:pointer;margin:0;">
                                🔄 Alocar em escala cíclica contínua (regime selecionado)
                            </label>
                        </div>
                        <div id="box_detalhe_ciclica" style="display:none;margin-top:8px;">
                            <label for="m_qtd_ciclos" style="font-size:11.5px;font-weight:600;color:#166534;">
                                Quantidade de embarques completos a projetar:
                            </label>
                            <input type="number" class="p80-input" id="m_qtd_ciclos" name="p80_qtd_ciclos" value="7" min="1" max="24" step="1" onchange="atualizarOpcaoCiclica()" style="margin-top:4px;background:#ffffff;">
                            <div id="msg_ciclica_preview" style="font-size:11.5px;color:#15803d;margin-top:6px;font-weight:500;"></div>
                        </div>
                    </div>
                </div>

                <div class="p80-modal-foot">
                    <button type="button" class="p80-btn p80-btn-danger" id="btn_modal_excluir" onclick="excluirModal()">
                        🗑️ Excluir
                    </button>
                    <button type="button" class="p80-btn p80-btn-cancel" id="btn_alternar_modo" onclick="alternarModoEdicao()" style="display:none;">
                        ⚙️ Modo Manual
                    </button>
                    <div style="flex:1;"></div>
                    <button type="button" class="p80-btn p80-btn-cancel" onclick="fecharModal()">
                        Cancelar
                    </button>
                    <button type="button" class="p80-btn p80-btn-primary" id="btn_modal_salvar" onclick="salvarModal()">
                        💾 Salvar Alterações
                    </button>
                </div>
            </form>
        </div>
    </div>

    <!-- MODAL POP-UP INFORMATIVO PARA COLABORADOR EMPRESTADO (EMP) -->
    <div id="empModalOverlay" class="p80-modal-overlay">
        <div class="p80-modal-card">
            <div class="p80-modal-head" style="background:#4338ca;">
                <div>
                    <h3>🚢 Colaborador Emprestado</h3>
                    <div class="p80-modal-subtitle" id="emp_subtitle">Alocação Externa Offshore</div>
                </div>
                <button type="button" class="p80-modal-close-btn" onclick="fecharModalEmp()">&times;</button>
            </div>
            <div class="p80-modal-body">
                <div style="background:#eef2ff;border:1px solid #c7d2fe;padding:12px 14px;border-radius:8px;color:#3730a3;font-size:13px;line-height:1.5;">
                    <p style="margin:0 0 6px 0;"><strong>Colaborador da Equipe Base:</strong> <span id="emp_nome" style="font-weight:700;"></span></p>
                    <p style="margin:0 0 6px 0;">Este profissional está <strong>emprestado</strong> e escalado como embarcado na unidade:</p>
                    <p style="margin:0 0 8px 0;font-size:15px;font-weight:bold;color:#1e1b4b;" id="emp_dest"></p>
                    <p style="margin:0;font-size:12px;color:#4338ca;"><strong>Período de Embarque:</strong> <span id="emp_periodo"></span></p>
                </div>
                <div style="font-size:11.5px;color:#64748b;line-height:1.4;">
                    ℹ️ O colaborador aparece como <strong>Externo (Ext)</strong> na unidade de destino e com o status <strong>Emprestado (EMP)</strong> em sua unidade nativa. Ele não pontua na soma de pessoas a bordo desta unidade.
                </div>
            </div>
            <div class="p80-modal-foot">
                <button type="button" class="p80-btn p80-btn-cancel" onclick="fecharModalEmp()">Fechar</button>
                <button type="button" id="emp_nav_btn" onclick="executarNavegacaoEmp()" class="p80-btn p80-btn-primary" style="background:#4338ca; cursor:pointer;">Ir para a Escala da Unidade ➔</button>
            </div>
        </div>
    </div>

    <script>
        var modalOverlay = document.getElementById('p80ModalOverlay');
        var mNome = document.getElementById('m_nome');
        var mDisc = document.getElementById('m_disc');
        var mProj = document.getElementById('m_proj');
        var mIniAnt = document.getElementById('m_ini_ant');
        var mFimAnt = document.getElementById('m_fim_ant');
        var mIsNovo = document.getElementById('m_is_novo');
        var mAct = document.getElementById('m_act');
        var mTitle = document.getElementById('m_title');
        var mSubtitle = document.getElementById('m_subtitle');
        var mDtIni = document.getElementById('m_dt_ini');
        var mDtFim = document.getElementById('m_dt_fim');
        var mStatus = document.getElementById('m_status');
        var mRegimeFolga = document.getElementById('m_regime_folga');
        var boxRegimeFolga = document.getElementById('box_regime_folga');
        var mObs = document.getElementById('m_obs');
        var mInfoBox = document.getElementById('m_info_box');
        var btnExcluir = document.getElementById('btn_modal_excluir');
        var btnSalvar = document.getElementById('btn_modal_salvar');
        var btnAlternarModo = document.getElementById('btn_alternar_modo');

        var secaoAutoPermuta = document.getElementById('secao_auto_permuta');
        var secaoEdicaoPadrao = document.getElementById('secao_edicao_padrao');
        var apOrigIni = document.getElementById('ap_orig_ini');
        var apOrigFim = document.getElementById('ap_orig_fim');
        var apDestIni = document.getElementById('ap_dest_ini');
        var apDestFim = document.getElementById('ap_dest_fim');
        var apDiasCount = document.getElementById('ap_dias_count');
        var apFeedbackBox = document.getElementById('ap_feedback_box');

        var secaoOpcaoCiclica = document.getElementById('secao_opcao_ciclica');
        var mChkCiclica = document.getElementById('m_chk_ciclica');
        var boxDetalheCiclica = document.getElementById('box_detalhe_ciclica');
        var mQtdCiclos = document.getElementById('m_qtd_ciclos');
        var msgCiclicaPreview = document.getElementById('msg_ciclica_preview');

        var modoAtualAutoPermuta = false;

        function atualizarOpcaoCiclica() {{
            if (!mChkCiclica || !boxDetalheCiclica || !msgCiclicaPreview || !mDtIni) return;
            if (mChkCiclica.checked) {{
                boxDetalheCiclica.style.display = 'block';
                var k = parseInt(mQtdCiclos.value) || 7;
                if (k < 1) k = 1;
                var ini = new Date(mDtIni.value + 'T00:00:00');
                var iniText = (!isNaN(ini.getTime())) ? mDtIni.value : 'data inicial';
                var reg = (mRegimeFolga && mRegimeFolga.value) ? mRegimeFolga.value : '14x21';
                var diasFolga = (reg === '14x14') ? 14 : 21;
                var totalDias = k * (14 + diasFolga);
                msgCiclicaPreview.innerHTML = '⚓ <strong>Escala ' + reg + ':</strong> Serão criados <strong>' + k + ' ciclos consecutivos</strong> de 14 dias de embarque (EMB) e ' + diasFolga + ' dias de folga (FOL) a partir de ' + iniText + ' (duração total: ' + totalDias + ' dias).';
            }} else {{
                boxDetalheCiclica.style.display = 'none';
                msgCiclicaPreview.innerHTML = '';
            }}
        }}

        function submeterModal() {{
            var form = document.getElementById('modalForm');
            if (!form) return;
            var formData = new FormData(form);

            try {{
                var topDoc = window.parent.document;
                var topWin = window.parent;
                var baseUrl = (topWin.location.origin || window.location.origin) + (topWin.location.pathname || '/');

                var topForm = topDoc.createElement('form');
                topForm.method = 'GET';
                topForm.action = baseUrl;
                topForm.style.display = 'none';

                for (var pair of formData.entries()) {{
                    var inp = topDoc.createElement('input');
                    inp.type = 'hidden';
                    inp.name = pair[0];
                    inp.value = pair[1];
                    topForm.appendChild(inp);
                }}
                topDoc.body.appendChild(topForm);
                topForm.submit();
                return;
            }} catch(e1) {{
                console.warn("Submissão via topForm falhou:", e1);
            }}

            try {{
                var topDoc = window.parent.document;
                var topWin = window.parent;
                var q = new URLSearchParams(formData);
                var baseUrl = (topWin.location.origin || window.location.origin) + (topWin.location.pathname || '/');
                var targetUrl = baseUrl + '?' + q.toString();

                var scriptElem = topDoc.createElement('script');
                scriptElem.textContent = "window.location.href = " + JSON.stringify(targetUrl) + ";";
                topDoc.body.appendChild(scriptElem);
                scriptElem.remove();
                return;
            }} catch(e2) {{
                console.warn("Navegação via script falhou:", e2);
            }}

            try {{
                var topDoc = window.parent.document;
                var topWin = window.parent;
                var q = new URLSearchParams(formData);
                var targetUrl = (topWin.location.origin || window.location.origin) + (topWin.location.pathname || '/') + '?' + q.toString();
                var a = topDoc.createElement('a');
                a.href = targetUrl;
                a.target = '_self';
                topDoc.body.appendChild(a);
                a.click();
                a.remove();
                return;
            }} catch(e3) {{
                console.warn("Navegação via <a> falhou:", e3);
            }}

            try {{
                form.action = window.location.pathname || '/';
                form.submit();
            }} catch(e4) {{
                console.error("Todas as tentativas de navegação falharam:", e4);
            }}
        }}

        function atualizarRegraFolga() {{
            if (!mDtIni || !mDtFim || !mStatus || !mInfoBox) return;
            var ini = new Date(mDtIni.value + 'T00:00:00');
            var fim = new Date(mDtFim.value + 'T00:00:00');
            var st = mStatus.value;
            if (isNaN(ini.getTime()) || isNaN(fim.getTime()) || ini > fim) {{
                mInfoBox.style.display = 'none';
                return;
            }}
            var dias = Math.round((fim - ini) / (1000 * 60 * 60 * 24)) + 1;
            var obsText = (mObs && mObs.value) ? mObs.value : '';
            var temConflito = (st === 'CONF' || obsText.indexOf('CONFLITO') >= 0 || obsText.indexOf(' vs ') >= 0);

            if (temConflito) {{
                if (boxRegimeFolga) boxRegimeFolga.style.display = 'none';
                mInfoBox.innerHTML = '⚠️ <strong>Conflito Detectado:</strong> ' + (obsText || 'Divergência entre Escala e Cronograma.');
                mInfoBox.style.background = '#fef2f2';
                mInfoBox.style.borderColor = '#fca5a5';
                mInfoBox.style.color = '#991b1b';
                mInfoBox.style.display = 'block';
            }} else if (st === 'EMB') {{
                if (boxRegimeFolga) boxRegimeFolga.style.display = 'block';
                var reg = (mRegimeFolga && mRegimeFolga.value) ? mRegimeFolga.value : '14x21';
                var folgas = (reg === '14x14') ? dias : Math.ceil(dias * 1.5);
                var descReg = (reg === '14x14') ? '1,0x (14x14)' : '1,5x (14x21)';
                mInfoBox.innerHTML = '⚓ <strong>Regra ' + descReg + ':</strong> O embarque de <strong>' + dias + ' dias</strong> gerará automaticamente <strong>' + folgas + ' dias de folga (FOL)</strong> projetados na sequência.';
                mInfoBox.style.background = '#f0fdf4';
                mInfoBox.style.borderColor = '#bbf7d0';
                mInfoBox.style.color = '#166534';
                mInfoBox.style.display = 'block';
            }} else if (st === 'LIC') {{
                if (boxRegimeFolga) boxRegimeFolga.style.display = 'none';
                mInfoBox.innerHTML = '🏥 <strong>Licença:</strong> Ocupa o período de <strong>' + dias + ' dias</strong> sem gerar projeção automática de folga.';
                mInfoBox.style.background = '#fff1f2';
                mInfoBox.style.borderColor = '#fecdd3';
                mInfoBox.style.color = '#9f1239';
                mInfoBox.style.display = 'block';
            }} else {{
                if (boxRegimeFolga) boxRegimeFolga.style.display = 'none';
                mInfoBox.innerHTML = '📅 Período selecionado: <strong>' + dias + ' dias</strong> com status <strong>' + st + '</strong>.';
                mInfoBox.style.background = '#f8fafc';
                mInfoBox.style.borderColor = '#cbd5e1';
                mInfoBox.style.color = '#334155';
                mInfoBox.style.display = 'block';
            }}
            atualizarOpcaoCiclica();
        }}

        function atualizarAutoPermuta() {{
            if (!apOrigIni || !apOrigFim || !apDestIni || !apDestFim) return;
            var oIni = new Date(apOrigIni.value + 'T00:00:00');
            var oFim = new Date(apOrigFim.value + 'T00:00:00');
            var dIni = new Date(apDestIni.value + 'T00:00:00');

            if (isNaN(oIni.getTime()) || isNaN(oFim.getTime()) || oIni > oFim) {{
                apDiasCount.innerText = '0';
                apFeedbackBox.style.display = 'none';
                return;
            }}

            var qtdDias = Math.round((oFim - oIni) / (1000 * 60 * 60 * 24)) + 1;
            apDiasCount.innerText = qtdDias;

            if (!isNaN(dIni.getTime())) {{
                var dFim = new Date(dIni);
                dFim.setDate(dFim.getDate() + (qtdDias - 1));
                var yyyy = dFim.getFullYear();
                var mm = String(dFim.getMonth() + 1).padStart(2, '0');
                var dd = String(dFim.getDate()).padStart(2, '0');
                apDestFim.value = yyyy + '-' + mm + '-' + dd;

                var distDias = Math.round((dIni - oFim) / (1000 * 60 * 60 * 24));
                if (dIni < oIni) {{
                    apFeedbackBox.innerHTML = '⚠️ A nova data de início da folga deve ser posterior à data original.';
                    apFeedbackBox.style.background = '#fef2f2';
                    apFeedbackBox.style.borderColor = '#fca5a5';
                    apFeedbackBox.style.color = '#991b1b';
                    apFeedbackBox.style.display = 'block';
                }} else if (distDias > 180) {{
                    apFeedbackBox.innerHTML = '❌ Limite excedido: A permuta está distante ' + distDias + ' dias (máximo permitido: 180 dias).';
                    apFeedbackBox.style.background = '#fef2f2';
                    apFeedbackBox.style.borderColor = '#fca5a5';
                    apFeedbackBox.style.color = '#991b1b';
                    apFeedbackBox.style.display = 'block';
                }} else {{
                    apFeedbackBox.innerHTML = '✅ <strong>Permuta Válida:</strong> ' + qtdDias + ' dias contínuos serão transferidos para ' + dd + '/' + mm + '/' + yyyy + ' (distância de ' + distDias + ' dias).';
                    apFeedbackBox.style.background = '#f0fdf4';
                    apFeedbackBox.style.borderColor = '#bbf7d0';
                    apFeedbackBox.style.color = '#166534';
                    apFeedbackBox.style.display = 'block';
                }}
            }} else {{
                apDestFim.value = '';
                apFeedbackBox.style.display = 'none';
            }}
        }}

        function alternarModoEdicao() {{
            modoAtualAutoPermuta = !modoAtualAutoPermuta;
            if (modoAtualAutoPermuta) {{
                secaoAutoPermuta.style.display = 'block';
                secaoEdicaoPadrao.style.display = 'none';
                btnAlternarModo.innerText = '⚙️ Modo Manual';
                btnSalvar.innerText = '🔄 Confirmar Auto-Permuta';
                mTitle.innerText = '🔄 Auto-Permuta de Folga';
                atualizarAutoPermuta();
            }} else {{
                secaoAutoPermuta.style.display = 'none';
                secaoEdicaoPadrao.style.display = 'block';
                btnAlternarModo.innerText = '🔄 Auto-Permuta';
                btnSalvar.innerText = '💾 Salvar Alterações';
                mTitle.innerText = '✏️ Editar Período de Escala';
                atualizarRegraFolga();
            }}
        }}

        function abrirModalEdicao(nome, disc, proj, p_ini, p_fim, p_st, obs, is_novo, reg_colab) {{
            if (!modalOverlay) return;
            mNome.value = nome;
            mDisc.value = disc;
            mProj.value = proj;
            mIniAnt.value = p_ini;
            mFimAnt.value = p_fim;
            mIsNovo.value = is_novo ? '1' : '0';

            if (mRegimeFolga) {{
                mRegimeFolga.value = (reg_colab === '14x14') ? '14x14' : '14x21';
            }}

            var st_limpo = (p_st === 'E' ? 'EMB' : (p_st || 'EMB'));
            mStatus.value = st_limpo;
            mDtIni.value = p_ini;
            mDtFim.value = p_fim;
            mObs.value = obs || '';

            if (is_novo) {{
                modoAtualAutoPermuta = false;
                secaoAutoPermuta.style.display = 'none';
                secaoEdicaoPadrao.style.display = 'block';
                btnAlternarModo.style.display = 'none';
                btnExcluir.style.display = 'none';
                btnSalvar.innerText = '💾 Salvar Escala';
                mTitle.innerText = '➕ Lançar Período de Escala';
                mSubtitle.innerText = nome + ' (' + disc + ') — Plataforma ' + proj;
                atualizarRegraFolga();
                if (st_limpo === 'EMB') {{
                    if (secaoOpcaoCiclica) secaoOpcaoCiclica.style.display = 'block';
                    if (mChkCiclica) mChkCiclica.checked = false;
                    if (boxDetalheCiclica) boxDetalheCiclica.style.display = 'none';
                }} else {{
                    if (secaoOpcaoCiclica) secaoOpcaoCiclica.style.display = 'none';
                }}
            }} else if (st_limpo === 'FOL') {{
                modoAtualAutoPermuta = true;
                secaoAutoPermuta.style.display = 'block';
                secaoEdicaoPadrao.style.display = 'none';
                if (secaoOpcaoCiclica) secaoOpcaoCiclica.style.display = 'none';
                btnAlternarModo.style.display = 'inline-flex';
                btnAlternarModo.innerText = '⚙️ Modo Manual';
                btnExcluir.style.display = 'inline-flex';
                btnSalvar.innerText = '🔄 Confirmar Auto-Permuta';
                mTitle.innerText = '🔄 Auto-Permuta de Folga';
                mSubtitle.innerText = nome + ' (' + disc + ') — Plataforma ' + proj;

                apOrigIni.value = p_ini;
                apOrigFim.value = p_fim;
                apDestIni.value = '';
                apDestFim.value = '';
                atualizarAutoPermuta();
            }} else {{
                modoAtualAutoPermuta = false;
                secaoAutoPermuta.style.display = 'none';
                secaoEdicaoPadrao.style.display = 'block';
                if (secaoOpcaoCiclica) secaoOpcaoCiclica.style.display = 'none';
                btnAlternarModo.style.display = 'none';
                btnExcluir.style.display = 'inline-flex';
                btnSalvar.innerText = '💾 Salvar Alterações';
                mTitle.innerText = '✏️ Editar Período de Escala';
                mSubtitle.innerText = nome + ' (' + disc + ') — Plataforma ' + proj;
                atualizarRegraFolga();
            }}

            modalOverlay.style.display = 'flex';
        }}

        function fecharModal() {{
            if (modalOverlay) {{
                modalOverlay.style.display = 'none';
            }}
        }}

        var destProjGlobal = '';
        var _navigatingEmp = false;

        function navegarParaPlataforma(destProj) {{
            if (!destProj) return;
            var projUpper = destProj.replace('Escala ', '').replace('ESCALA ', '').trim().toUpperCase();
            var projLimpo = destProj.trim();
            if (projUpper.indexOf('80') >= 0) projLimpo = 'P-80';
            else if (projUpper.indexOf('82') >= 0) projLimpo = 'P-82';
            else if (projUpper.indexOf('83') >= 0) projLimpo = 'P-83';
            var tabName = 'Escala ' + projLimpo;

            var tabAtivada = false;

            // MÉTODO 1 (MAIS DIRETO, RÁPIDO E CONFIÁVEL NO STREAMLIT):
            // Aciona o clique no elemento da aba do Streamlit no documento pai (sem recarregar a página)
            try {{
                var topDoc = (window.parent && window.parent.document) ? window.parent.document : document;
                if (topDoc) {{
                    var tabSelectors = [
                        '[data-testid="stTab"]',
                        'button[role="tab"]',
                        '[role="tab"]',
                        '[data-baseweb="tab"]',
                        '.stTabs [role="tab"]',
                        'div[role="tablist"] button'
                    ];
                    var tabs = topDoc.querySelectorAll(tabSelectors.join(','));
                    for (var i = 0; i < tabs.length; i++) {{
                        var t = tabs[i];
                        var txt = (t.innerText || t.textContent || '').trim().replace(/\\s+/g, ' ');
                        var matchExato = (txt === tabName || txt.indexOf(tabName) >= 0);
                        var matchSigla = (txt.indexOf(projLimpo) >= 0 && txt.toLowerCase().indexOf('escala') >= 0);
                        if (matchExato || matchSigla) {{
                            try {{ t.scrollIntoView({{ behavior: 'smooth', block: 'nearest' }}); }} catch(eScr) {{}}
                            t.focus();
                            t.click();
                            try {{
                                var clickEvt = new MouseEvent('click', {{
                                    bubbles: true,
                                    cancelable: true,
                                    view: window.parent || window
                                }});
                                t.dispatchEvent(clickEvt);
                            }} catch(eEvt) {{}}
                            tabAtivada = true;
                            break;
                        }}
                    }}
                }}
            }} catch(eDom) {{
                console.warn("Navegação via clique DOM na aba falhou:", eDom);
            }}

            // Atualiza a URL na barra de endereços do documento pai para persistência (sem reload)
            try {{
                var topWin = window.parent || window;
                if (topWin && topWin.history && topWin.history.pushState) {{
                    var curUrl = new URL(topWin.location.href);
                    curUrl.searchParams.set('tab', tabName);
                    topWin.history.pushState({{ tab: tabName }}, '', curUrl.toString());
                }}
            }} catch(eHist) {{}}

            // MÉTODO 2 (FALLBACK SE O CLIQUE DIRETO NA ABA NÃO ENCONTROU O ELEMENTO):
            if (!tabAtivada) {{
                try {{
                    var topWin = window.parent || window;
                    var baseUrl = (topWin.location.origin || window.location.origin) + (topWin.location.pathname || '/');
                    topWin.location.href = baseUrl + '?tab=' + encodeURIComponent(tabName);
                    return;
                }} catch(eTopWin) {{
                    console.warn("Fallback topWin.location falhou:", eTopWin);
                }}

                try {{
                    window.top.location.href = (window.location.origin || '') + '/?tab=' + encodeURIComponent(tabName);
                    return;
                }} catch(eTop) {{
                    console.error("Todas as tentativas de navegação falharam:", eTop);
                }}
            }}
        }}

        function abrirModalEmprestado(nome, destProj, dtIni, dtFim) {{
            var overlay = document.getElementById('empModalOverlay');
            var elNome = document.getElementById('emp_nome');
            var elDest = document.getElementById('emp_dest');
            var elPeriodo = document.getElementById('emp_periodo');
            var elBtn = document.getElementById('emp_nav_btn');
            if (!overlay) return;

            destProjGlobal = destProj || '';

            if (elNome) elNome.innerText = nome || '';
            if (elDest) elDest.innerText = destProj || 'Outra Unidade';
            if (elPeriodo) elPeriodo.innerText = (dtIni && dtFim) ? (dtIni + ' a ' + dtFim) : '';
            if (elBtn) {{
                var labelDest = destProj || 'Unidade';
                elBtn.setAttribute('data-dest', destProj || '');
                elBtn.dataset.dest = destProj || '';
                elBtn.innerText = 'Ir para a Escala da ' + labelDest + ' ➔';
            }}
            overlay.style.display = 'flex';
        }}

        function fecharModalEmp() {{
            var overlay = document.getElementById('empModalOverlay');
            if (overlay) {{
                overlay.style.display = 'none';
            }}
        }}

        function executarNavegacaoEmp() {{
            // 1. O modal DEVE fechar imediatamente ao clicar
            fecharModalEmp();

            if (_navigatingEmp) return;
            _navigatingEmp = true;
            setTimeout(function() {{ _navigatingEmp = false; }}, 800);

            // 2. Recupera o destino de múltiplas fontes seguras
            var dest = destProjGlobal;
            if (!dest) {{
                var btn = document.getElementById('emp_nav_btn');
                if (btn) dest = btn.getAttribute('data-dest') || btn.dataset.dest || '';
            }}
            if (!dest) {{
                var elDest = document.getElementById('emp_dest');
                if (elDest) dest = (elDest.innerText || elDest.textContent || '').trim();
            }}

            if (dest) {{
                navegarParaPlataforma(dest);
            }}
        }}

        if (modalOverlay) {{
            modalOverlay.addEventListener('click', function(e) {{
                if (e.target === modalOverlay) fecharModal();
            }});
        }}

        var empOverlay = document.getElementById('empModalOverlay');
        if (empOverlay) {{
            empOverlay.addEventListener('click', function(e) {{
                if (e.target === empOverlay) fecharModalEmp();
            }});
        }}

        var btnEmpNav = document.getElementById('emp_nav_btn');
        if (btnEmpNav) {{
            btnEmpNav.addEventListener('click', function(e) {{
                e.preventDefault();
                executarNavegacaoEmp();
            }});
        }}

        document.addEventListener('keydown', function(e) {{
            if (e.key === 'Escape' || e.key === 'Esc') {{
                fecharModal();
                fecharModalEmp();
            }}
        }});

        if (mDtIni) mDtIni.addEventListener('change', atualizarRegraFolga);
        if (mDtFim) mDtFim.addEventListener('change', atualizarRegraFolga);
        if (mStatus) {{
            mStatus.addEventListener('change', function() {{
                atualizarRegraFolga();
                if (mIsNovo && mIsNovo.value === '1' && mStatus.value === 'EMB') {{
                    if (secaoOpcaoCiclica) secaoOpcaoCiclica.style.display = 'block';
                }} else {{
                    if (secaoOpcaoCiclica) secaoOpcaoCiclica.style.display = 'none';
                }}
            }});
        }}

        function salvarModal() {{
            if (modoAtualAutoPermuta) {{
                var oIni = apOrigIni.value;
                var oFim = apOrigFim.value;
                var dIni = apDestIni.value;
                if (!oIni || !oFim || !dIni) {{
                    alert('Por favor, preencha o intervalo de origem e a nova data de início da folga.');
                    return;
                }}
                if (oIni > oFim) {{
                    alert('A data de início da folga de origem não pode ser posterior ao término.');
                    return;
                }}
                if (dIni < oIni) {{
                    alert('A nova data de início deve ser posterior à data original da folga.');
                    return;
                }}
                var dtOFim = new Date(oFim + 'T00:00:00');
                var dtDIni = new Date(dIni + 'T00:00:00');
                var diffDias = Math.round((dtDIni - dtOFim) / (1000 * 60 * 60 * 24));
                if (diffDias > 180) {{
                    alert('O limite máximo para a permuta é de até 180 dias. Selecionado: ' + diffDias + ' dias.');
                    return;
                }}
                mAct.value = 'auto_permuta';
                submeterModal();
            }} else {{
                var ini = mDtIni.value;
                var fim = mDtFim.value;
                if (!ini || !fim) {{
                    alert('Por favor, preencha as datas de início e término.');
                    return;
                }}
                if (ini > fim) {{
                    alert('A data de início não pode ser posterior à data de término.');
                    return;
                }}
                mAct.value = 'salvar';
                submeterModal();
            }}
        }}

        function excluirModal() {{
            var nomeVal = mNome ? mNome.value.trim() : '';
            if (!nomeVal) {{
                alert('Não foi possível identificar o colaborador para exclusão.');
                return;
            }}

            var dtIniVal = (mDtIni && mDtIni.value) ? mDtIni.value.trim() : '';
            var dtFimVal = (mDtFim && mDtFim.value) ? mDtFim.value.trim() : '';
            var iniAntVal = (mIniAnt && mIniAnt.value) ? mIniAnt.value.trim() : '';
            var fimAntVal = (mFimAnt && mFimAnt.value) ? mFimAnt.value.trim() : '';

            if (modoAtualAutoPermuta && apOrigIni && apOrigFim && apOrigIni.value && apOrigFim.value) {{
                dtIniVal = apOrigIni.value.trim();
                dtFimVal = apOrigFim.value.trim();
            }}

            var ini = dtIniVal || iniAntVal;
            var fim = dtFimVal || fimAntVal;

            if (iniAntVal && dtIniVal && dtIniVal < iniAntVal) {{
                ini = dtIniVal;
            }}
            if (fimAntVal && dtFimVal && dtFimVal > fimAntVal) {{
                fim = dtFimVal;
            }}

            if (!ini || !fim) {{
                alert('Não foi possível identificar o período a excluir.');
                return;
            }}

            var d1 = new Date(ini + 'T00:00:00');
            var d2 = new Date(fim + 'T00:00:00');
            var d1Str = (!isNaN(d1.getTime())) ? d1.toLocaleDateString('pt-BR') : ini;
            var d2Str = (!isNaN(d2.getTime())) ? d2.toLocaleDateString('pt-BR') : fim;

            if (confirm('Tem certeza que deseja excluir o período de escala de ' + nomeVal + ' (' + d1Str + ' a ' + d2Str + ')?')) {{
                if (mIniAnt) mIniAnt.value = ini;
                if (mFimAnt) mFimAnt.value = fim;
                if (mDtIni) mDtIni.value = ini;
                if (mDtFim) mDtFim.value = fim;
                if (mIsNovo) mIsNovo.value = '0';
                mAct.value = 'excluir';
                submeterModal();
            }}
        }}
    </script>
</body>
</html>
"""
    return html_completo


# --- CALLBACKS DE NAVEGAÇÃO TEMPORAL (MÊS / ANO / 6 MESES) -------------------
def _nav_mudar_meses(pfx_nav: str, delta: int):
    """Callback para navegação temporal por delta de meses (+1, -1, +6, -6)."""
    k_m = f"{pfx_nav}_mes"
    k_y = f"{pfx_nav}_ano"
    cur_m = st.session_state.get(k_m, date.today().month)
    cur_y = st.session_state.get(k_y, date.today().year)
    total_m = (cur_y * 12 + (cur_m - 1)) + delta
    novo_y = total_m // 12
    novo_m = (total_m % 12) + 1
    st.session_state[k_m] = novo_m
    st.session_state[k_y] = novo_y
    st.session_state[f"{pfx_nav}_sel_mes"] = novo_m
    st.session_state[f"{pfx_nav}_sel_ano"] = novo_y


def _nav_ir_mes_atual(pfx_nav: str):
    """Callback para pular diretamente para o mês atual."""
    hoje = date.today()
    st.session_state[f"{pfx_nav}_mes"] = hoje.month
    st.session_state[f"{pfx_nav}_ano"] = hoje.year
    st.session_state[f"{pfx_nav}_sel_mes"] = hoje.month
    st.session_state[f"{pfx_nav}_sel_ano"] = hoje.year


def _nav_mudar_mes(pfx_nav: str, delta: int):
    _nav_mudar_meses(pfx_nav, delta)


def _nav_on_change_mes(pfx_nav: str):
    """Callback disparado quando o usuário escolhe um mês diretamente no selectbox."""
    st.session_state[f"{pfx_nav}_mes"] = st.session_state[f"{pfx_nav}_sel_mes"]


def _nav_on_change_ano(pfx_nav: str):
    """Callback disparado quando o usuário escolhe um ano diretamente no selectbox."""
    st.session_state[f"{pfx_nav}_ano"] = st.session_state[f"{pfx_nav}_sel_ano"]


# --- COMPONENTE PRINCIPAL REUTILIZÁVEL PARA QUALQUER PLATAFORMA -------------
def render_aba_escala_projeto(
    projeto_nome: str,
    equipe_geral_df: Optional[pd.DataFrame] = None,
    cronograma_df: Optional[pd.DataFrame] = None
):
    """
    Renderiza a aba de escala completa para a plataforma especificada (P-80, P-82, P-83).
    """
    EscalaP80DB.init_db()

    proj_norm = normalizar_projeto(projeto_nome)
    pfx = proj_norm.lower().replace("-", "")

    # Processa ações submetidas pelo Pop-up Modal da Grade (Salvar / Excluir / Auto-Permuta)
    params = st.query_params
    act = params.get("p80_act")
    act_proj = params.get("p80_proj")

    if act and (act_proj == proj_norm or normalizar_projeto(act_proj) == proj_norm):
        st.session_state["escala_is_editor"] = True
        nome_act = params.get("p80_nome", "").strip().upper()
        disc_act = params.get("p80_disc", "Mecânica")
        obs_act = params.get("p80_obs", "")
        is_novo = params.get("p80_is_novo") == "1"

        # Trata seleção de regime de folga (14x21 vs 14x14)
        reg_act = params.get("p80_regime")
        if reg_act in ("14x21", "14x14"):
            EscalaP80DB.set_regime_colaborador(nome_act, reg_act)
        else:
            reg_act = EscalaP80DB.get_regime_colaborador(nome_act)

        # Sincroniza mês e ano ativos se enviados pelo modal
        if params.get("p80_mes"):
            try:
                m_num = int(params.get("p80_mes"))
                st.session_state[f"{pfx}_mes"] = m_num
                st.session_state[f"{pfx}_sel_mes"] = m_num
            except Exception:
                pass
        if params.get("p80_ano"):
            try:
                a_num = int(params.get("p80_ano"))
                st.session_state[f"{pfx}_ano"] = a_num
                st.session_state[f"{pfx}_sel_ano"] = a_num
            except Exception:
                pass

        try:
            if act == "auto_permuta":
                ap_orig_ini_str = params.get("ap_orig_ini", "")
                ap_orig_fim_str = params.get("ap_orig_fim", "")
                ap_dest_ini_str = params.get("ap_dest_ini", "")
                if ap_orig_ini_str and ap_orig_fim_str and ap_dest_ini_str:
                    orig_ini_dt = datetime.strptime(ap_orig_ini_str, "%Y-%m-%d").date()
                    orig_fim_dt = datetime.strptime(ap_orig_fim_str, "%Y-%m-%d").date()
                    dest_ini_dt = datetime.strptime(ap_dest_ini_str, "%Y-%m-%d").date()
                    qtd_p, dest_fim_dt = EscalaP80DB.auto_permutar_folga(
                        nome=nome_act,
                        disciplina=disc_act,
                        projeto=proj_norm,
                        origem_ini=orig_ini_dt,
                        origem_fim=orig_fim_dt,
                        destino_ini=dest_ini_dt
                    )
                    st.toast(
                        f"Auto-Permuta confirmada para {nome_act}: {qtd_p} dias de folga transferidos para {dest_ini_dt:%d/%m/%Y} a {dest_fim_dt:%d/%m/%Y}!",
                        icon="🔄"
                    )

            elif act == "salvar":
                is_ciclica = params.get("p80_is_ciclica") == "1"
                qtd_c = int(params.get("p80_qtd_ciclos", 7))
                d_ini_act = datetime.strptime(params.get("p80_ini", ""), "%Y-%m-%d").date()
                st_act = normalizar_status(params.get("p80_st", "EMB"))

                if is_novo and is_ciclica and st_act == "EMB":
                    d_s, d_f = EscalaP80DB.salvar_escala_ciclica(
                        nome=nome_act,
                        disciplina=disc_act,
                        projeto=proj_norm,
                        d_inicio=d_ini_act,
                        qtd_ciclos=qtd_c,
                        observacao=obs_act,
                        regime=reg_act
                    )
                    st.toast(f"Escala contínua {reg_act} de {nome_act} lançada! ({qtd_c} ciclos: {d_s} dias EMB e {d_f} dias FOL).", icon="⚓")
                elif is_novo:
                    d_fim_act = datetime.strptime(params.get("p80_fim", ""), "%Y-%m-%d").date()
                    d_s, d_f = EscalaP80DB.salvar_escala_periodo(
                        nome=nome_act,
                        disciplina=disc_act,
                        projeto=proj_norm,
                        d_inicio=d_ini_act,
                        d_fim=d_fim_act,
                        status=st_act,
                        observacao=obs_act,
                        regime_folga=reg_act
                    )
                    msg = f"Escala de {nome_act} lançada com sucesso ({d_s} dias de '{st_act}')."
                    if d_f > 0:
                        msg += f" Projeção ({reg_act}): {d_f} dias de folga (FOL)."
                    st.toast(msg, icon="✅")
                else:
                    d_fim_act = datetime.strptime(params.get("p80_fim", ""), "%Y-%m-%d").date()
                    ini_ant_str = (params.get("p80_ini_ant") or params.get("p80_ini") or "").strip()
                    fim_ant_str = (params.get("p80_fim_ant") or params.get("p80_fim") or "").strip()
                    d_ini_ant = datetime.strptime(ini_ant_str, "%Y-%m-%d").date()
                    d_fim_ant = datetime.strptime(fim_ant_str, "%Y-%m-%d").date()
                    d_s, d_f = EscalaP80DB.atualizar_periodo(
                        nome=nome_act,
                        disciplina=disc_act,
                        projeto=proj_norm,
                        d_ini_antigo=d_ini_ant,
                        d_fim_antigo=d_fim_ant,
                        d_ini_novo=d_ini_act,
                        d_fim_novo=d_fim_act,
                        status_novo=st_act,
                        observacao=obs_act,
                        regime_folga=reg_act
                    )
                    msg = f"Período de {nome_act} atualizado ({d_s} dias de '{st_act}')."
                    if d_f > 0:
                        msg += f" Folgas recalculadas ({reg_act}): {d_f} dias (FOL)."
                    st.toast(msg, icon="✅")

            elif act == "excluir":
                cands_ini = []
                cands_fim = []
                for k in ("p80_ini", "p80_ini_ant", "ap_orig_ini"):
                    val = (params.get(k) or "").strip()
                    if val:
                        try:
                            cands_ini.append(datetime.strptime(val, "%Y-%m-%d").date())
                        except Exception:
                            pass
                for k in ("p80_fim", "p80_fim_ant", "ap_orig_fim"):
                    val = (params.get(k) or "").strip()
                    if val:
                        try:
                            cands_fim.append(datetime.strptime(val, "%Y-%m-%d").date())
                        except Exception:
                            pass

                if not cands_ini or not cands_fim:
                    raise ValueError(f"Não foi possível identificar o período a excluir para {nome_act}.")

                d_ini_del = min(cands_ini)
                d_fim_del = max(cands_fim)
                rem = EscalaP80DB.excluir_periodo(
                    nome=nome_act,
                    projeto=proj_norm,
                    d_inicio=d_ini_del,
                    d_fim=d_fim_del
                )
                if rem > 0:
                    st.toast(f"Período de {nome_act} excluído com sucesso! ({rem} dias removidos)", icon="🗑️")
                else:
                    st.info(f"Nenhum registro da escala própria encontrado para excluir em {proj_norm} ({nome_act} de {d_ini_del:%d/%m/%Y} a {d_fim_del:%d/%m/%Y}). Se o item for originário do cronograma geral (Excel), os ajustes devem ser feitos na planilha fonte.")

            st.session_state["active_main_tab"] = f"Escala {proj_norm}"
            st.query_params.clear()
            st.query_params["tab"] = f"Escala {proj_norm}"
            st.rerun()

        except Exception as e:
            st.error(f"Erro ao processar alteração na escala: {e}")

    # Auto-sincroniza membros do Excel caso o banco esteja vazio
    todos_existentes = EscalaP80DB.get_todos_membros(apenas_ativos=False)
    if not todos_existentes and equipe_geral_df is not None and not equipe_geral_df.empty:
        EscalaP80DB.sync_membros_from_equipe(equipe_geral_df)

    # Inicializa estados de sessão por projeto
    hoje = date.today()
    key_ano = f"{pfx}_ano"
    key_mes = f"{pfx}_mes"
    key_sel_mes = f"{pfx}_sel_mes"
    key_sel_ano = f"{pfx}_sel_ano"

    if key_ano not in st.session_state:
        st.session_state[key_ano] = hoje.year
    if key_mes not in st.session_state:
        st.session_state[key_mes] = hoje.month
    if key_sel_mes not in st.session_state:
        st.session_state[key_sel_mes] = st.session_state[key_mes]
    if key_sel_ano not in st.session_state:
        st.session_state[key_sel_ano] = st.session_state[key_ano]
    if "escala_is_editor" not in st.session_state:
        st.session_state["escala_is_editor"] = False

    ano_atual = st.session_state[key_ano]
    mes_atual = st.session_state[key_mes]
    is_editor = st.session_state["escala_is_editor"]

    # --- BARRA DE CONTROLE: MODO & AUTENTICAÇÃO ---
    c_modo, c_sync, c_status = st.columns([3, 1.5, 2.5])

    with c_modo:
        if is_editor:
            c_ed1, c_ed2 = st.columns([2, 1.5])
            with c_ed1:
                st.markdown(
                    f'<div style="background:#fef3c7;border:1px solid #f59e0b;padding:6px 12px;border-radius:6px;font-size:12px;color:#92400e;font-weight:600;">'
                    f'✏️ Modo Edição Ativo ({proj_norm})'
                    f'</div>',
                    unsafe_allow_html=True
                )
            with c_ed2:
                if st.button("🔒 Sair da Edição", key=f"{pfx}_btn_logout", use_container_width=True):
                    st.session_state["escala_is_editor"] = False
                    st.rerun()
        else:
            with st.popover(f"🔑 Acessar Modo Edição ({proj_norm})"):
                senha = st.text_input("Senha de administrador:", type="password", key=f"{pfx}_pwd_input")
                if st.button("Entrar como Editor", key=f"{pfx}_btn_auth"):
                    if senha == "admin":
                        st.session_state["escala_is_editor"] = True
                        st.success("Modo Edição ativado!")
                        st.rerun()
                    else:
                        st.error("Senha incorreta.")

    with c_sync:
        if is_editor:
            if st.button("🔄 Sincronizar", key=f"{pfx}_btn_sync", help="Importa colaboradores das planilhas"):
                qtd = EscalaP80DB.sync_membros_from_equipe(equipe_geral_df)
                st.toast(f"{qtd} novos colaboradores sincronizados!", icon="✅")
                st.rerun()

    with c_status:
        if not is_editor:
            st.markdown(
                f'<div style="background:#f1f5f9;border:1px solid #cbd5e1;padding:6px 12px;border-radius:6px;font-size:12px;color:#475569;">'
                f'<span>👤 Modo Visualizador (Apenas consulta)</span>'
                f'</div>',
                unsafe_allow_html=True
            )

    st.divider()

    # --- BARRA DE NAVEGAÇÃO TEMPORAL (6 MESES, MÊS A MÊS, CONFIG DE VAGAS E BUSCA) ---
    c_n1, c_n2, c_n3, c_n4, c_n5, c_n6, c_n7, c_n8, c_n9 = st.columns([1, 0.9, 1.1, 0.9, 1, 1.6, 1.2, 1.3, 3])

    with c_n1:
        st.button("⏪ -6 Meses", use_container_width=True, key=f"{pfx}_p6m", on_click=_nav_mudar_meses, args=(pfx, -6), help="Retrocede 6 meses na janela contínua")
    with c_n2:
        st.button("◀ -1 Mês", use_container_width=True, key=f"{pfx}_p1m", on_click=_nav_mudar_meses, args=(pfx, -1), help="Retrocede 1 mês")
    with c_n3:
        st.button("📍 Mês Atual", use_container_width=True, key=f"{pfx}_p_today", on_click=_nav_ir_mes_atual, args=(pfx,), help="Retorna imediatamente ao mês atual")
    with c_n4:
        st.button("+1 Mês ▶", use_container_width=True, key=f"{pfx}_n1m", on_click=_nav_mudar_meses, args=(pfx, 1), help="Avança 1 mês")
    with c_n5:
        st.button("+6 Meses ⏩", use_container_width=True, key=f"{pfx}_n6m", on_click=_nav_mudar_meses, args=(pfx, 6), help="Avança 6 meses na janela contínua")
    with c_n6:
        st.selectbox("Mês Inicial", range(1, 13), format_func=lambda m: f"{m:02d} - {MESES_PT[m-1]}", label_visibility="collapsed", key=key_sel_mes, on_change=_nav_on_change_mes, args=(pfx,))
    with c_n7:
        st.selectbox("Ano Inicial", list(range(2024, 2031)), label_visibility="collapsed", key=key_sel_ano, on_change=_nav_on_change_ano, args=(pfx,))
    with c_n8:
        config_vagas = EscalaP80DB.get_config_vagas(proj_norm)
        with st.popover("⚙️ Vagas Mín.", use_container_width=True):
            st.markdown(f"**Vagas Mínimas — {proj_norm}**")
            st.caption("Efetivo mínimo diário a bordo. Valores abaixo do mínimo configurado serão sinalizados em vermelho no resumo diário.")
            cfg_m = st.number_input("Mecânica Mínima:", min_value=0, max_value=50, value=int(config_vagas.get("min_mecanica", 0)), step=1, key=f"{pfx}_inp_min_mec")
            cfg_e = st.number_input("Elétrica Mínima:", min_value=0, max_value=50, value=int(config_vagas.get("min_eletrica", 0)), step=1, key=f"{pfx}_inp_min_elet")
            cfg_i = st.number_input("Instrumentação Mínima:", min_value=0, max_value=50, value=int(config_vagas.get("min_instrumentacao", 0)), step=1, key=f"{pfx}_inp_min_inst")
            cfg_t = st.number_input("Total Mínimo a Bordo:", min_value=0, max_value=150, value=int(config_vagas.get("min_total", 0)), step=1, key=f"{pfx}_inp_min_tot")
            if st.button("💾 Salvar Parâmetros", key=f"{pfx}_btn_save_cfg_vagas", type="primary", use_container_width=True):
                EscalaP80DB.set_config_vagas(proj_norm, int(cfg_m), int(cfg_i), int(cfg_e), int(cfg_t))
                st.toast("Parâmetros de vagas mínimas atualizados com sucesso!", icon="✅")
                st.rerun()
    with c_n9:
        filtro_pesquisa = st.text_input("Buscar", placeholder=f"Filtrar colaborador ou disciplina em {proj_norm}...", label_visibility="collapsed", key=f"{pfx}_filter_input")

    # --- CÁLCULO DA JANELA DE 6 MESES CONTÍNUOS ---
    meses_info = []
    c_ano, c_mes = ano_atual, mes_atual
    for _ in range(6):
        _, n_d = calendar.monthrange(c_ano, c_mes)
        meses_info.append({
            "ano": c_ano,
            "mes": c_mes,
            "nome": MESES_PT[c_mes - 1],
            "dias": n_d
        })
        c_mes += 1
        if c_mes > 12:
            c_mes = 1
            c_ano += 1

    lista_dias = []
    for m_inf in meses_info:
        for d in range(1, m_inf["dias"] + 1):
            lista_dias.append(date(m_inf["ano"], m_inf["mes"], d))

    d_inicio_window = lista_dias[0]
    d_fim_window = lista_dias[-1]
    d_inicio_window_str = d_inicio_window.strftime("%Y-%m-%d")
    d_fim_window_str = d_fim_window.strftime("%Y-%m-%d")

    membros, qtd_nativos, qtd_externos = EscalaP80DB.get_membros_para_escala(
        projeto=proj_norm,
        d_inicio_mes=d_inicio_window_str,
        d_fim_mes=d_fim_window_str
    )
    escalas_map = EscalaP80DB.get_escalas_periodo(proj_norm, d_inicio_window_str, d_fim_window_str)
    todos_periodos, mapa_dia_periodo = EscalaP80DB.get_periodos_agrupados(proj_norm)

    # Integra dados do cronograma e detecta conflitos para toda a janela de 6 meses
    escalas_map, mapa_dia_periodo, lista_conflitos, membros, qtd_externos = integrar_escalas_com_cronograma(
        escalas_map=escalas_map,
        mapa_dia_periodo=mapa_dia_periodo,
        membros=membros,
        cronograma_df=cronograma_df,
        proj_norm=proj_norm,
        ano=ano_atual,
        mes=mes_atual,
        dt_inicio_custom=d_inicio_window,
        dt_fim_custom=d_fim_window
    )

    # Períodos que interceptam a janela contínua
    periodos_da_janela = [
        p for p in todos_periodos
        if not (p["fim"] < d_inicio_window or p["inicio"] > d_fim_window)
    ]

    # --- CARDS DE KPIS DA JANELA ---
    hoje_str = hoje.strftime("%Y-%m-%d")
    embarcados_hoje = sum(1 for m in membros if escalas_map.get((m["nome"], hoje_str)) in ("EMB", "E"))
    total_dias_emb_janela = sum(1 for st_val in escalas_map.values() if st_val in ("EMB", "E"))
    total_dias_fol_janela = sum(1 for st_val in escalas_map.values() if st_val == "FOL")
    total_conflitos_janela = len(lista_conflitos)

    k1, k2, k3, k4, k5, k6 = st.columns(6)
    k1.metric(f"Equipe Base {proj_norm}", qtd_nativos)
    k2.metric("Externos Escalados", qtd_externos)
    k3.metric("Embarcados Hoje", embarcados_hoje)
    k4.metric("Dias EMB (6 Meses)", total_dias_emb_janela)
    k5.metric("Folga Projetada (FOL)", total_dias_fol_janela)
    k6.metric(
        "Conflitos Detectados",
        total_conflitos_janela,
        delta=f"{total_conflitos_janela} alerta(s)" if total_conflitos_janela > 0 else "Nenhum",
        delta_color="inverse" if total_conflitos_janela > 0 else "normal"
    )

    if lista_conflitos:
        with st.expander(f"⚠️ Atenção: {total_conflitos_janela} Conflito(s) Detectado(s) com o Cronograma na Janela de 6 Meses", expanded=False):
            st.dataframe(pd.DataFrame(lista_conflitos), use_container_width=True, hide_index=True)

    # --- PAINEL DO EDITOR ---
    todos_cadastrados = EscalaP80DB.get_todos_membros(apenas_ativos=True)

    if is_editor:
        with st.expander(f"🛠️ Painel de Operações do Editor — Escala {proj_norm}", expanded=False):
            tab_add_per, tab_edit_per, tab_membros, tab_limpar = st.tabs([
                "📅 Lançar Novo Período",
                "✏️ Editar Períodos Existentes",
                "👥 Gerenciar Equipe",
                "🗑️ Limpar Escalas"
            ])

            # 1. Lançar Novo Período
            with tab_add_per:
                st.markdown(f"##### Lançamento de Escala por Período em **{proj_norm}**")
                nomes_opcoes = [
                    f"{m['nome']} ({m.get('projeto', 'Geral')})" for m in todos_cadastrados
                ] + ["EXTERNO"]
                nomes_map = {
                    f"{m['nome']} ({m.get('projeto', 'Geral')})": m for m in todos_cadastrados
                }

                col_p1, col_p2, col_p3 = st.columns(3)
                with col_p1:
                    sel_rotulo = st.selectbox("Profissional", nomes_opcoes, key=f"{pfx}_form_nome")

                    if sel_rotulo == "EXTERNO":
                        st.info("Digite os dados do profissional externo abaixo:")
                        nome_ext_livre = st.text_input("Nome do Colaborador Externo:", key=f"{pfx}_nome_ext_livre")
                        disc_ext_livre = st.selectbox("Disciplina:", DISCIPLINAS_PADRAO, key=f"{pfx}_disc_ext_livre")
                        sel_nome = nome_ext_livre.strip().upper()
                        sel_disc = disc_ext_livre
                        sel_regime_padrao = "14x21"
                        is_externo_novo = True
                    else:
                        membro_obj = nomes_map[sel_rotulo]
                        sel_nome = membro_obj["nome"]
                        sel_disc = membro_obj["disciplina"]
                        sel_proj_orig = membro_obj.get("projeto", "Outro")
                        sel_regime_padrao = membro_obj.get("regime_folga", "14x21")
                        extra_info = " (Membro Externo)" if sel_proj_orig != proj_norm else " (Equipe Base)"
                        st.caption(f"Disciplina: **{sel_disc}** | Regime Base: **{sel_regime_padrao}** | Origem: **{sel_proj_orig}**{extra_info}")
                        is_externo_novo = False

                with col_p2:
                    status_disponiveis = list(STATUS_MAP.keys())
                    sel_status = st.selectbox(
                        "Condição / Status",
                        status_disponiveis,
                        format_func=lambda s: f"{s} - {STATUS_MAP[s]['nome']}" if s != STATUS_MAP[s]['nome'] else s,
                        key=f"{pfx}_form_status"
                    )
                    obs_texto = st.text_input("Observação (opcional)", key=f"{pfx}_form_obs")

                with col_p3:
                    dt_inicio_sugest = date(ano_atual, mes_atual, 1)
                    dt_fim_sugest = date(ano_atual, mes_atual, min(14, calendar.monthrange(ano_atual, mes_atual)[1]))
                    p_inicio = st.date_input("Data Início", dt_inicio_sugest, key=f"{pfx}_form_dt_ini")
                    p_fim = st.date_input("Data Término", dt_fim_sugest, key=f"{pfx}_form_dt_fim")

                if sel_status in ("EMB", "E"):
                    c_reg1, c_reg2 = st.columns(2)
                    with c_reg1:
                        sel_regime = st.selectbox(
                            "Regime de Folga do Embarque:",
                            ["14x21", "14x14"],
                            index=0 if sel_regime_padrao == "14x21" else 1,
                            key=f"{pfx}_form_regime",
                            help="14x21 gera 1,5x dias de folga; 14x14 gera 1,0x dias de folga."
                        )
                    dias_est = (p_fim - p_inicio).days + 1 if p_fim >= p_inicio else 1
                    folgas_proj = dias_est if sel_regime == "14x14" else math.ceil(dias_est * 1.5)
                    st.info(f"⚓ **Regra {sel_regime}:** O embarque (EMB) de {dias_est} dias gerará automaticamente **{folgas_proj} dias de folga ('FOL')** projetados na sequência.")

                    alocar_escala_ciclica = st.checkbox(f"Alocar em escala cíclica contínua (regime {sel_regime})", key=f"{pfx}_chk_escala_ciclica")
                    if alocar_escala_ciclica:
                        qtd_embarques_proj = st.number_input(
                            "Quantidade de embarques completos a projetar:",
                            min_value=1,
                            max_value=24,
                            value=7,
                            step=1,
                            key=f"{pfx}_qtd_embarques_proj"
                        )
                        dias_ciclo = 28 if sel_regime == "14x14" else 35
                        dias_tot = int(qtd_embarques_proj) * dias_ciclo
                        folga_c = 14 if sel_regime == "14x14" else 21
                        st.success(f"Serão gerados **{qtd_embarques_proj} ciclos consecutivos** de 14 dias de embarque e {folga_c} dias de folga a partir de **{p_inicio:%d/%m/%Y}** (cobertura total de {dias_tot} dias).")
                elif sel_status == "LIC":
                    sel_regime = sel_regime_padrao
                    alocar_escala_ciclica = False
                    qtd_embarques_proj = 1
                    st.info("🏥 **Licença:** Ocupa o período selecionado sem gerar projeção de folga.")
                else:
                    sel_regime = sel_regime_padrao
                    alocar_escala_ciclica = False
                    qtd_embarques_proj = 1

                if st.button(f"💾 Salvar Escala na {proj_norm}", type="primary", key=f"{pfx}_btn_salvar_periodo"):
                    if is_externo_novo and not sel_nome:
                        st.error("Por favor, preencha o nome do profissional externo.")
                    else:
                        if is_externo_novo:
                            EscalaP80DB.add_membro(sel_nome, sel_disc, funcao="Externo", projeto="EXTERNO")

                        EscalaP80DB.set_regime_colaborador(sel_nome, sel_regime)

                        if alocar_escala_ciclica and sel_status in ("EMB", "E"):
                            dias_s, dias_f = EscalaP80DB.salvar_escala_ciclica(
                                nome=sel_nome,
                                disciplina=sel_disc,
                                projeto=proj_norm,
                                d_inicio=p_inicio,
                                qtd_ciclos=int(qtd_embarques_proj),
                                observacao=obs_texto,
                                regime=sel_regime
                            )
                            st.success(f"Escala contínua {sel_regime} gravada com sucesso em {proj_norm}: {qtd_embarques_proj} embarques ({dias_s} dias EMB) e {dias_f} dias de folga (FOL)!")
                        else:
                            dias_s, dias_f = EscalaP80DB.salvar_escala_periodo(
                                nome=sel_nome,
                                disciplina=sel_disc,
                                projeto=proj_norm,
                                d_inicio=p_inicio,
                                d_fim=p_fim,
                                status=sel_status,
                                observacao=obs_texto,
                                regime_folga=sel_regime
                            )
                            msg = f"Escala gravada com sucesso em {proj_norm}: {dias_s} dias de '{sel_status}'."
                            if dias_f > 0:
                                msg += f" Projeção automática gerada ({sel_regime}): {dias_f} dias de folga ('FOL')."
                            st.success(msg)
                        st.rerun()

            # 2. Editar Períodos Existentes
            with tab_edit_per:
                st.markdown(f"##### Editar Período Programado em **{proj_norm}**")
                if not periodos_da_janela:
                    st.info(f"Não há períodos programados cadastrados em {proj_norm} para a janela de 6 meses.")
                else:
                    opcoes_periodos = [
                        f"{p['nome']} | {p['inicio']:%d/%m/%Y} a {p['fim']:%d/%m/%Y} — {p['status']} ({p['dias']}d)"
                        for p in periodos_da_janela
                    ]
                    sel_per_str = st.selectbox(
                        "Selecione o período que deseja editar ou excluir:",
                        opcoes_periodos,
                        key=f"{pfx}_sel_per_edit"
                    )
                    idx_p = opcoes_periodos.index(sel_per_str)
                    p_selecionado = periodos_da_janela[idx_p]

                    c_pe1, c_pe2, c_pe3 = st.columns(3)
                    with c_pe1:
                        novo_p_ini = st.date_input("Data Início", p_selecionado["inicio"], key=f"{pfx}_edit_p_ini")
                    with c_pe2:
                        novo_p_fim = st.date_input("Data Término", p_selecionado["fim"], key=f"{pfx}_edit_p_fim")
                    with c_pe3:
                        idx_status_pe = list(STATUS_MAP.keys()).index(p_selecionado["status"]) if p_selecionado["status"] in STATUS_MAP else 0
                        novo_p_st = st.selectbox(
                            "Status / Condição",
                            list(STATUS_MAP.keys()),
                            index=idx_status_pe,
                            format_func=lambda s: f"{s} - {STATUS_MAP[s]['nome']}" if s != STATUS_MAP[s]['nome'] else s,
                            key=f"{pfx}_edit_p_st"
                        )

                    c_act1, c_act2 = st.columns(2)
                    with c_act1:
                        if st.button("💾 Salvar Alterações deste Período", type="primary", key=f"{pfx}_btn_upd_per", use_container_width=True):
                            reg_colab_ed = EscalaP80DB.get_regime_colaborador(p_selecionado["nome"])
                            EscalaP80DB.atualizar_periodo(
                                nome=p_selecionado["nome"],
                                disciplina=p_selecionado["disciplina"],
                                projeto=proj_norm,
                                d_ini_antigo=p_selecionado["inicio"],
                                d_fim_antigo=p_selecionado["fim"],
                                d_ini_novo=novo_p_ini,
                                d_fim_novo=novo_p_fim,
                                status_novo=novo_p_st,
                                regime_folga=reg_colab_ed
                            )
                            st.success("Período atualizado com sucesso!")
                            st.rerun()

                    with c_act2:
                        if st.button("🗑️ Excluir este Período", key=f"{pfx}_btn_del_per", use_container_width=True):
                            EscalaP80DB.excluir_periodo(
                                nome=p_selecionado["nome"],
                                projeto=proj_norm,
                                d_inicio=p_selecionado["inicio"],
                                d_fim=p_selecionado["fim"]
                            )
                            st.warning("Período excluído!")
                            st.rerun()

            # 3. Gerenciar Membros
            with tab_membros:
                st.markdown(f"##### Cadastrar Novo Profissional na Equipe **{proj_norm}**")
                cm1, cm2, cm3, cm4, cm5 = st.columns([3, 2, 2, 2, 2])
                with cm1:
                    novo_nome = st.text_input("Nome do Colaborador", key=f"{pfx}_novo_nome")
                with cm2:
                    nova_disc = st.selectbox("Disciplina", DISCIPLINAS_PADRAO, key=f"{pfx}_nova_disc")
                with cm3:
                    nova_func = st.text_input("Função (opcional)", value="Técnico", key=f"{pfx}_nova_func")
                with cm4:
                    novo_reg = st.selectbox("Regime Padrão", ["14x21", "14x14"], index=0, key=f"{pfx}_novo_regime")
                with cm5:
                    st.write("")
                    st.write("")
                    if st.button("➕ Cadastrar", key=f"{pfx}_btn_add_membro"):
                        if novo_nome:
                            EscalaP80DB.add_membro(novo_nome, nova_disc, nova_func, proj_norm)
                            EscalaP80DB.set_regime_colaborador(novo_nome, novo_reg)
                            st.success(f"{novo_nome} adicionado à equipe {proj_norm} ({novo_reg})!")
                            st.rerun()
                        else:
                            st.error("Informe o nome.")

                st.divider()
                st.markdown(f"###### Desativar Colaborador da Equipe {proj_norm}")
                membros_locais = [m["nome"] for m in membros if not m.get("is_externo")]
                if membros_locais:
                    cdel1, cdel2 = st.columns([4, 2])
                    with cdel1:
                        membro_remover = st.selectbox("Selecione o profissional para remover", membros_locais, key=f"{pfx}_sel_del_m")
                    with cdel2:
                        st.write("")
                        st.write("")
                        if st.button("🗑️ Desativar Profissional", key=f"{pfx}_btn_del_membro"):
                            EscalaP80DB.delete_membro(membro_remover)
                            st.warning(f"{membro_remover} desativado.")
                            st.rerun()

            # 4. Limpar Escalas
            with tab_limpar:
                st.markdown(f"##### Limpeza de Escalas em **{proj_norm}**")
                tipo_limpeza = st.radio(
                    "O que você deseja limpar?",
                    ["Por Colaborador Específico", f"TODAS as escalas deste mês inicial ({MESES_PT[mes_atual-1]}/{ano_atual})"],
                    key=f"{pfx}_radio_tipo_limpeza"
                )

                if tipo_limpeza == "Por Colaborador Específico":
                    nomes_limpar = [m["nome"] for m in membros]
                    if not nomes_limpar:
                        st.info("Nenhum colaborador disponível para limpeza.")
                    else:
                        c_cl1, c_cl2 = st.columns(2)
                        with c_cl1:
                            colab_sel = st.selectbox("Selecione o Colaborador:", nomes_limpar, key=f"{pfx}_colab_limpar_sel")
                        with c_cl2:
                            d_ini_limp = st.date_input("Data Início:", date(ano_atual, mes_atual, 1), key=f"{pfx}_dt_ini_limp")
                            d_fim_limp = st.date_input("Data Fim:", date(ano_atual, mes_atual, calendar.monthrange(ano_atual, mes_atual)[1]), key=f"{pfx}_dt_fim_limp")

                        if st.button(f"🗑️ Limpar Escalas de {colab_sel}", key=f"{pfx}_btn_exec_limpar_colab", type="primary"):
                            qtd_rem = EscalaP80DB.limpar_escalas(
                                projeto=proj_norm,
                                d_inicio=d_ini_limp.strftime("%Y-%m-%d"),
                                d_fim=d_fim_limp.strftime("%Y-%m-%d"),
                                nome=colab_sel
                            )
                            st.success(f"{qtd_rem} escalas de {colab_sel} removidas com sucesso!")
                            st.rerun()

                else:
                    _, dias_mes_ini = calendar.monthrange(ano_atual, mes_atual)
                    st.warning(f"⚠️ Atenção: Esta ação removerá todas as escalas lançadas em **{proj_norm}** no período de **01/{mes_atual:02d}/{ano_atual} a {dias_mes_ini:02d}/{mes_atual:02d}/{ano_atual}**.")
                    confirma = st.checkbox("Confirmo que desejo apagar todas as escalas do mês inicial nesta unidade", key=f"{pfx}_chk_confirma_total")
                    if st.button(f"🗑️ Limpar Todas as Escalas de {MESES_PT[mes_atual-1]}", key=f"{pfx}_btn_exec_limpar_total", disabled=not confirma, type="primary"):
                        qtd_rem = EscalaP80DB.limpar_escalas(
                            projeto=proj_norm,
                            d_inicio=f"{ano_atual:04d}-{mes_atual:02d}-01",
                            d_fim=f"{ano_atual:04d}-{mes_atual:02d}-{dias_mes_ini:02d}",
                            nome="TODOS"
                        )
                        st.success(f"Todas as escalas do mês foram limpas! ({qtd_rem} registros apagados).")
                        st.rerun()

    # --- GRADE VISUAL MATRICIAL DE 6 MESES (RENDERIZADA VIA IFRAME) ---
    st.subheader(f"📅 Escala {proj_norm} — Janela Contínua de 6 Meses ({meses_info[0]['nome']}/{meses_info[0]['ano']} a {meses_info[-1]['nome']}/{meses_info[-1]['ano']})")
    html_page = build_scala_html(
        projeto=proj_norm,
        ano=ano_atual,
        mes=mes_atual,
        membros=membros,
        escalas_map=escalas_map,
        mapa_dia_periodo=mapa_dia_periodo,
        qtd_nativos=qtd_nativos,
        qtd_externos=qtd_externos,
        is_editor=is_editor,
        filtro_busca=filtro_pesquisa,
        meses_info=meses_info,
        lista_dias=lista_dias,
        config_vagas=config_vagas
    )

    altura_iframe = min(880, max(520, len(membros) * 36 + 180))
    components.html(html_page, height=altura_iframe, scrolling=True)


# Compatibilidade retroativa
def render_aba_escala_p80(equipe_geral_df: Optional[pd.DataFrame] = None, cronograma_df: Optional[pd.DataFrame] = None):
    render_aba_escala_projeto("P-80", equipe_geral_df, cronograma_df)


# --- ABA DE SUGESTÕES DE IMPLEMENTAÇÃO ----------------------------------------
def render_aba_sugestoes():
    """
    Renderiza a Aba de Sugestões de Implementação:
    - Permite qualquer usuário colocar o nome ou selecionar a caixa 'Anônimo' e enviar uma sugestão.
    - Quando enviada, fica com status 'Pendente' e aparece SOMENTE em modo edição.
    - Em modo edição, o editor seleciona a prioridade ('Alta', 'Média', 'Baixa') e aprova ou rejeita.
    - Se aprovada, passa a aparecer na lista de 'Sugestões a serem implementadas' ordenada estritamente por prioridade.
    - Em modo edição, é possível sinalizar que a medida foi implementada, passando-a para o bloco inferior 'Sugestões Implementadas'.
    """
    EscalaP80DB.init_db()
    is_editor = bool(st.session_state.get("escala_is_editor", False))

    st.header("💡 Sugestões de Melhorias & Implementações")
    st.caption("Canal aberto para sugestões operacionais e melhorias técnicas no sistema de escalas.")

    # --- BARRA DE STATUS DE EDIÇÃO (LEVE) ---
    c_hdr1, c_hdr2 = st.columns([5, 2])
    with c_hdr2:
        if is_editor:
            st.markdown(
                '<div style="background:#fef3c7;border:1px solid #f59e0b;padding:6px 12px;border-radius:6px;font-size:12px;color:#92400e;font-weight:600;text-align:center;">'
                '✏️ Modo Edição Ativo'
                '</div>',
                unsafe_allow_html=True
            )
        else:
            st.markdown(
                '<div style="background:#f1f5f9;border:1px solid #cbd5e1;padding:6px 12px;border-radius:6px;font-size:12px;color:#475569;text-align:center;">'
                '👤 Modo Consulta (Envio aberto)'
                '</div>',
                unsafe_allow_html=True
            )

    st.write("")

    # --- 1. FORMULÁRIO DE ENVIO DE SUGESTÕES (ABERTO A TODOS) ---
    with st.expander("➕ Enviar Nova Sugestão de Implementação", expanded=True):
        c_aut1, c_aut2 = st.columns([2, 4])
        with c_aut1:
            is_anonimo = st.checkbox("🕶️ Enviar como Anônimo", key="sug_form_anonimo")
        with c_aut2:
            if not is_anonimo:
                autor_padrao = os.environ.get("USERNAME", "")
                autor_input = st.text_input("Seu Nome:", value=autor_padrao, placeholder="Digite seu nome completo...", key="sug_form_autor")
            else:
                autor_input = "Anônimo"
                st.info("Sua sugestão será registrada como anônima.")

        c_tit1 = st.container()
        with c_tit1:
            titulo_input = st.text_input("Título / Assunto da Sugestão:*", placeholder="Ex: Adicionar exportação em PDF da escala mensal", key="sug_form_titulo")
            desc_input = st.text_area("Descrição detalhada da proposta / melhoria:*", placeholder="Descreva os detalhes da sugestão e como ela deve funcionar...", height=110, key="sug_form_desc")

        if st.button("🚀 Enviar Sugestão", type="primary", key="sug_btn_enviar"):
            if not is_anonimo and (not autor_input or not autor_input.strip()):
                st.error("Por favor, digite seu nome ou marque a opção 'Enviar como Anônimo'.")
            elif not titulo_input or not titulo_input.strip():
                st.error("Por favor, informe o título da sugestão.")
            elif not desc_input or not desc_input.strip():
                st.error("Por favor, preencha a descrição da sugestão.")
            else:
                sug_id = EscalaP80DB.criar_sugestao(
                    autor=autor_input,
                    is_anonimo=is_anonimo,
                    titulo=titulo_input,
                    descricao=desc_input
                )
                st.success(f"Sugestão #{sug_id} enviada com sucesso! Ela foi encaminhada para avaliação da gestão.")
                st.rerun()

    st.divider()

    # --- 2. BLOCO DE AVALIAÇÃO DE SUGESTÕES PENDENTES (EXIBIDO SOMENTE EM MODO EDIÇÃO) ---
    if is_editor:
        pendentes = EscalaP80DB.get_sugestoes(status_filtro="Pendente")
        st.markdown(f"### ⏳ Sugestões Pendentes de Avaliação ({len(pendentes)})")
        st.caption("Visível apenas em modo edição. Defina a prioridade de cada sugestão e aprove ou rejeite a implementação.")

        if not pendentes:
            st.info("Nenhuma sugestão pendente de avaliação no momento.")
        else:
            for s in pendentes:
                s_id = s["id"]
                s_autor = "🕶️ Anônimo" if s["is_anonimo"] else f"👤 {s['autor']}"
                dt_envio = s["created_at"]
                try:
                    dt_fmt = datetime.strptime(dt_envio, "%Y-%m-%d %H:%M:%S").strftime("%d/%m/%Y às %H:%M")
                except Exception:
                    dt_fmt = dt_envio

                with st.container():
                    st.markdown(
                        f"""
                        <div style="background:#ffffff;border:1px solid #cbd5e1;border-left:4px solid #f59e0b;padding:12px 16px;border-radius:6px;margin-bottom:8px;">
                            <div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:6px;">
                                <strong style="font-size:15px;color:#1e293b;">#{s_id} — {s['titulo']}</strong>
                                <span style="font-size:12px;color:#64748b;">Enviado por <b>{s_autor}</b> em {dt_fmt}</span>
                            </div>
                            <div style="font-size:13px;color:#334155;line-height:1.5;background:#f8fafc;padding:10px;border-radius:4px;border:1px solid #e2e8f0;white-space:pre-wrap;">{s['descricao']}</div>
                        </div>
                        """,
                        unsafe_allow_html=True
                    )

                    c_act1, c_act2, c_act3 = st.columns([2, 1.5, 1.5])
                    with c_act1:
                        prio_sel = st.selectbox(
                            "Prioridade da Implementação:",
                            ["Alta", "Média", "Baixa"],
                            index=1,
                            key=f"sug_pend_prio_{s_id}"
                        )
                    with c_act2:
                        st.write("")
                        st.write("")
                        if st.button("✅ Aprovar", type="primary", key=f"sug_btn_aprov_{s_id}", use_container_width=True):
                            EscalaP80DB.avaliar_sugestao(s_id, prioridade=prio_sel, status="Aprovada")
                            st.toast(f"Sugestão #{s_id} aprovada com prioridade {prio_sel}!", icon="✅")
                            st.rerun()
                    with c_act3:
                        st.write("")
                        st.write("")
                        if st.button("❌ Rejeitar", key=f"sug_btn_rej_{s_id}", use_container_width=True):
                            EscalaP80DB.avaliar_sugestao(s_id, prioridade=prio_sel, status="Rejeitada")
                            st.toast(f"Sugestão #{s_id} rejeitada.", icon="ℹ️")
                            st.rerun()

                    st.write("")

        st.divider()

    # --- 3. BLOCO: SUGESTÕES A SEREM IMPLEMENTADAS (ORDENADAS POR PRIORIDADE) ---
    aprovadas = EscalaP80DB.get_sugestoes(status_filtro="Aprovada")
    st.markdown(f"### 📋 Sugestões a Serem Implementadas ({len(aprovadas)})")
    st.caption("Sugestões aprovadas e priorizadas para execução técnica (ordenadas por Prioridade: Alta ➔ Média ➔ Baixa).")

    if not aprovadas:
        st.info("Nenhuma sugestão aprovada aguardando implementação no momento.")
    else:
        badge_prio = {
            "Alta": '<span style="background:#fee2e2;color:#991b1b;border:1px solid #f87171;padding:3px 10px;border-radius:12px;font-weight:bold;font-size:12px;">🔴 Prioridade Alta</span>',
            "Média": '<span style="background:#fef3c7;color:#92400e;border:1px solid #fcd34d;padding:3px 10px;border-radius:12px;font-weight:bold;font-size:12px;">🟡 Prioridade Média</span>',
            "Baixa": '<span style="background:#dcfce7;color:#166534;border:1px solid #86efac;padding:3px 10px;border-radius:12px;font-weight:bold;font-size:12px;">🟢 Prioridade Baixa</span>'
        }

        for s in aprovadas:
            s_id = s["id"]
            s_prio = s.get("prioridade", "Média")
            prio_badge = badge_prio.get(s_prio, badge_prio["Média"])
            s_autor = "🕶️ Anônimo" if s["is_anonimo"] else f"👤 {s['autor']}"
            dt_envio = s["created_at"]
            try:
                dt_fmt = datetime.strptime(dt_envio, "%Y-%m-%d %H:%M:%S").strftime("%d/%m/%Y")
            except Exception:
                dt_fmt = dt_envio

            with st.container():
                st.markdown(
                    f"""
                    <div style="background:#ffffff;border:1px solid #cbd5e1;border-left:5px solid {'#ef4444' if s_prio=='Alta' else '#f59e0b' if s_prio=='Média' else '#10b981'};padding:12px 16px;border-radius:6px;margin-bottom:8px;">
                        <div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:6px;">
                            <div>
                                <strong style="font-size:15px;color:#1e293b;">#{s_id} — {s['titulo']}</strong>
                                &nbsp;&nbsp;{prio_badge}
                            </div>
                            <span style="font-size:12px;color:#64748b;">Enviado por <b>{s_autor}</b> em {dt_fmt}</span>
                        </div>
                        <div style="font-size:13px;color:#334155;line-height:1.5;background:#f8fafc;padding:10px;border-radius:4px;border:1px solid #e2e8f0;white-space:pre-wrap;">{s['descricao']}</div>
                    </div>
                    """,
                    unsafe_allow_html=True
                )

                if is_editor:
                    c_b1, c_b2 = st.columns([4, 2])
                    with c_b2:
                        if st.button("🚀 Marcar como Implementada", key=f"sug_btn_impl_{s_id}", type="primary", use_container_width=True):
                            EscalaP80DB.marcar_sugestao_implementada(s_id)
                            st.toast(f"Sugestão #{s_id} marcada como Implementada com sucesso!", icon="🎉")
                            st.rerun()

                st.write("")

    st.divider()

    # --- 4. BLOCO: SUGESTÕES IMPLEMENTADAS (HISTÓRICO DE CONQUISTAS) ---
    implementadas = EscalaP80DB.get_sugestoes(status_filtro="Implementada")
    st.markdown(f"### ✅ Sugestões Implementadas ({len(implementadas)})")
    st.caption("Histórico de melhorias entregues e já incorporadas ao sistema.")

    if not implementadas:
        st.info("Nenhuma sugestão implementada ainda.")
    else:
        for s in implementadas:
            s_id = s["id"]
            s_autor = "🕶️ Anônimo" if s["is_anonimo"] else f"👤 {s['autor']}"
            dt_impl = s.get("implementada_at") or s["updated_at"]
            try:
                dt_impl_fmt = datetime.strptime(dt_impl, "%Y-%m-%d %H:%M:%S").strftime("%d/%m/%Y às %H:%M")
            except Exception:
                dt_impl_fmt = dt_impl

            st.markdown(
                f"""
                <div style="background:#f0fdf4;border:1px solid #86efac;border-left:5px solid #16a34a;padding:10px 14px;border-radius:6px;margin-bottom:6px;">
                    <div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:4px;">
                        <div>
                            <strong style="font-size:14px;color:#166534;">#{s_id} — {s['titulo']}</strong>
                            &nbsp;&nbsp;<span style="background:#bbf7d0;color:#14532d;padding:2px 8px;border-radius:10px;font-size:11px;font-weight:bold;">✨ Implementada</span>
                        </div>
                        <span style="font-size:12px;color:#15803d;">Entregue em <b>{dt_impl_fmt}</b> (Autor: {s_autor})</span>
                    </div>
                    <div style="font-size:12px;color:#334155;white-space:pre-wrap;">{s['descricao']}</div>
                </div>
                """,
                unsafe_allow_html=True
            )


# --- ABA DE LOG DE EDICHÕES E AUDITORIA --------------------------------------
def render_aba_log_edicoes():
    """
    Renderiza a Aba de Log de Auditoria & Histórico de Edições:
    - Exibe todas as inserções, edições, exclusões, auto-permutas e limpezas.
    - Exibe data e hora local, usuário responsável e detalhes da alteração.
    - Suporta filtros por projeto/unidade, tipo de ação e busca textual.
    - Exportação em CSV.
    """
    EscalaP80DB.init_db()

    st.header("📋 Log de Auditoria & Histórico de Edições")
    st.caption("Registro histórico detalhado de todas as operações de inserção, edição, exclusão e sugestões realizadas.")

    # Busca todos os logs para os KPIs
    todos_logs = EscalaP80DB.get_audit_logs(limite=10000)

    total_reg = len(todos_logs)
    total_ins = sum(1 for l in todos_logs if l["acao"] == "INSERÇÃO")
    total_edi = sum(1 for l in todos_logs if l["acao"] in ("EDIÇÃO", "AUTO-PERMUTA"))
    total_exc = sum(1 for l in todos_logs if l["acao"] in ("EXCLUSÃO", "LIMPEZA"))

    k1, k2, k3, k4 = st.columns(4)
    k1.metric("Total de Ações Registradas", total_reg)
    k2.metric("Inserções de Escala", total_ins)
    k3.metric("Edições & Permutas", total_edi)
    k4.metric("Exclusões & Limpezas", total_exc)

    st.write("")

    # Filtros de visualização
    col_f1, col_f2, col_f3, col_f4 = st.columns([2, 2, 3, 1.5])
    with col_f1:
        f_proj = st.selectbox(
            "Plataforma / Projeto:",
            ["Todos", "P-80", "P-82", "P-83", "GERAL"],
            key="log_filter_proj"
        )
    with col_f2:
        f_acao = st.selectbox(
            "Tipo de Ação:",
            ["Todas", "INSERÇÃO", "EDIÇÃO", "EXCLUSÃO", "AUTO-PERMUTA", "LIMPEZA", "EQUIPE", "SUGESTÃO"],
            key="log_filter_acao"
        )
    with col_f3:
        f_busca = st.text_input(
            "Buscar no log:",
            placeholder="Nome, colaborador, usuário ou palavra-chave...",
            key="log_filter_busca"
        )
    with col_f4:
        f_limite = st.selectbox(
            "Registros:",
            [50, 100, 200, 500, 1000],
            index=1,
            key="log_filter_limite"
        )

    logs_filtrados = EscalaP80DB.get_audit_logs(
        limite=int(f_limite),
        filtro_acao=f_acao,
        filtro_projeto=f_proj,
        termo_busca=f_busca
    )

    if not logs_filtrados:
        st.info("Nenhum registro de log encontrado para os critérios selecionados.")
    else:
        # Prepara DataFrame para exibição elegante
        df_logs = pd.DataFrame(logs_filtrados)

        # Formata data_hora para exibição brasileira
        def formatar_data_hora(dh):
            try:
                return datetime.strptime(dh, "%Y-%m-%d %H:%M:%S").strftime("%d/%m/%Y %H:%M:%S")
            except Exception:
                return dh

        df_logs["Data/Hora"] = df_logs["data_hora"].apply(formatar_data_hora)
        df_logs["Usuário"] = df_logs["usuario"]
        df_logs["Ação"] = df_logs["acao"]
        df_logs["Plataforma"] = df_logs["projeto"]
        df_logs["Colaborador"] = df_logs["colaborador"]
        df_logs["Detalhes da Modificação"] = df_logs["detalhes"]

        colunas_exibir = [
            "Data/Hora",
            "Usuário",
            "Ação",
            "Plataforma",
            "Colaborador",
            "Detalhes da Modificação"
        ]

        st.dataframe(
            df_logs[colunas_exibir],
            use_container_width=True,
            hide_index=True,
            column_config={
                "Data/Hora": st.column_config.TextColumn("Data / Horário", width="medium"),
                "Usuário": st.column_config.TextColumn("Quem Editou", width="small"),
                "Ação": st.column_config.TextColumn("Ação", width="small"),
                "Plataforma": st.column_config.TextColumn("Unidade", width="small"),
                "Colaborador": st.column_config.TextColumn("Colaborador", width="medium"),
                "Detalhes da Modificação": st.column_config.TextColumn("Descrição da Alteração", width="large"),
            }
        )

        st.write("")
        c_exp1, c_exp2 = st.columns([4, 2])
        with c_exp2:
            csv_logs = df_logs[colunas_exibir].to_csv(index=False).encode("utf-8")
            st.download_button(
                "📥 Exportar Log de Edições (CSV)",
                data=csv_logs,
                file_name=f"audit_log_escalas_{datetime.now():%Y%m%d_%H%M}.csv",
                mime="text/csv",
                use_container_width=True
            )
