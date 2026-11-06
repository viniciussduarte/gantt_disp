"""
Componentes de interface de usuário Streamlit: sidebar, uploads, KPIs, conflitos, equipe e exportação (v0.3).
"""

from datetime import date, timedelta
from typing import Any, Dict, List, Optional, Set, Tuple
import pandas as pd
import streamlit as st

from .config import Config


def render_sidebar_filters(
    equipe_df: pd.DataFrame,
    hoje: date,
) -> Tuple[List[str], List[str], List[str], pd.Timestamp, pd.Timestamp, bool, Dict[str, Any]]:
    """
    Renderiza os controles interativos da barra lateral:
    - Botão de limpeza de cache
    - Upload opcional de planilhas para simulação
    - Filtros múltiplos por Disciplina, Projeto e Função
    - Seleção de período temporal (presets ou personalizado)
    - Toggle de filtro de disponibilidade exclusiva

    Retorna:
    (sel_discs, sel_projs, sel_funcoes, d_inicio, d_fim, only_available, uploaded_files_dict)
    """
    with st.sidebar:
        st.header("Filtros")
        if st.button("🔄 Atualizar dados", use_container_width=True, help="Recarrega as planilhas ignorando o cache"):
            st.cache_data.clear()
            st.rerun()

        # Controle de Modo Edição (Editor vs Visualizador)
        is_editor = bool(st.session_state.get("escala_is_editor", False))
        if is_editor:
            st.success("🔒 Modo Edição Ativo")
            if st.button("🚪 Sair do Modo Edição", key="sidebar_btn_logout", use_container_width=True):
                st.session_state["escala_is_editor"] = False
                st.rerun()
        else:
            with st.popover("🔒 Acessar Modo Edição"):
                senha = st.text_input("Senha de administrador:", type="password", key="sidebar_pwd_input")
                if st.button("Entrar", key="sidebar_btn_login", use_container_width=True):
                    if senha == "admin":
                        st.session_state["escala_is_editor"] = True
                        st.success("Modo Edição ativado!")
                        st.rerun()
                    else:
                        st.error("Senha incorreta.")

        # Upload opcional de arquivos para simulação ou substituição temporária
        with st.expander("📁 Upload de Planilhas (Opcional)", expanded=False):
            st.caption("Substitua temporariamente as planilhas padrão para simular novos cenários:")
            up_estaleiro = st.file_uploader(
                "Planejamento Estaleiro (.xlsx)",
                type=["xlsx"],
                key="uploader_estaleiro",
            )
            up_ferias = st.file_uploader(
                "Férias (.xlsx)",
                type=["xlsx"],
                key="uploader_ferias",
            )
            up_geral = st.file_uploader(
                "Planejamento Geral (.xlsx)",
                type=["xlsx"],
                key="uploader_geral",
            )

        uploaded_files = {
            "estaleiro": up_estaleiro,
            "ferias": up_ferias,
            "geral": up_geral,
        }

        all_discs = sorted([x for x in equipe_df["Disciplina"].dropna().unique()])
        sel_discs = st.multiselect("Disciplina", all_discs, default=all_discs)

        all_projs = sorted([x for x in equipe_df["Projeto"].dropna().unique()])
        sel_projs = st.multiselect("Projeto", all_projs, default=all_projs)

        funcao_map = {"SUPERVISOR": "LIDERANÇA", "COORDENADOR": "LIDERANÇA"}
        equipe_df["FunçãoAgrupada"] = equipe_df["Função"].map(
            lambda f: funcao_map.get(f, f) if pd.notna(f) else f
        )
        all_funcoes = sorted([x for x in equipe_df["FunçãoAgrupada"].dropna().unique()])
        sel_funcoes = st.multiselect("Função", all_funcoes, default=all_funcoes)

        preset = st.selectbox("Período", list(Config.DATE_PRESETS.keys()), index=1)
        dias = Config.DATE_PRESETS[preset]
        if dias is not None:
            d_inicio = pd.Timestamp(hoje - timedelta(days=7))
            d_fim = pd.Timestamp(hoje + timedelta(days=dias))
            st.caption(f"{d_inicio:%d/%m/%Y} a {d_fim:%d/%m/%Y}")
        else:
            col1, col2 = st.columns(2)
            d_inicio = pd.Timestamp(col1.date_input("Início", hoje - timedelta(days=7)))
            d_fim = pd.Timestamp(col2.date_input("Fim", hoje + timedelta(days=90)))

        if d_inicio > d_fim:
            st.sidebar.warning("Data de início posterior à data de fim. Invertendo...")
            d_inicio, d_fim = d_fim, d_inicio

        only_available = st.checkbox(
            "Apenas disponíveis no período",
            help="Mostra apenas quem não tem nenhum evento agendado dentro do período selecionado",
        )

    return sel_discs, sel_projs, sel_funcoes, d_inicio, d_fim, only_available, uploaded_files


def render_kpi_cards(
    total_filtrados: int,
    ativos_hoje: int,
    ferias_hoje: int,
    total_conflitos: int,
) -> None:
    """
    Renderiza os 4 cards de KPIs principais no topo:
    - Total de Colaboradores Filtrados.
    - Total de Colaboradores Ativos na Data de Hoje.
    - Total em Férias na Data de Hoje.
    - Total de Conflitos Detectados (com destaque visual caso maior que zero).
    """
    k1, k2, k3, k4 = st.columns(4)
    k1.metric("Colaboradores Filtrados", total_filtrados)
    k2.metric("Ativos Hoje", ativos_hoje, help="Colaboradores com atividade presencial, embarque ou estaleiro no dia atual")
    k3.metric("Em Férias Hoje", ferias_hoje, help="Colaboradores com férias vigentes na data de hoje")

    if total_conflitos > 0:
        k4.metric(
            "⚠️ Conflitos Detectados",
            f"{total_conflitos} conflito(s)",
            delta=f"+{total_conflitos} sobreposições",
            delta_color="inverse",
            help="Conflitos de agenda detectados. Clique na aba 'Conflitos' para detalhes.",
        )
    else:
        k4.metric("Conflitos Detectados", "0", help="Nenhuma colisão de agenda encontrada")


def render_conflicts_tab(conflicts_df: pd.DataFrame, hoje: date) -> None:
    """
    Renderiza a aba de conflitos de agenda com destaque de severidade (Crítico vs Operacional),
    divisão por plataforma e download do relatório de sobreposições em CSV/Excel.
    """
    st.subheader("⚠️ Relatório de Conflitos de Agenda")
    st.caption(f"Detecção vetorizada de sobreposições com data de término posterior a {hoje:%d/%m/%Y}.")

    if conflicts_df.empty:
        st.success("✅ Nenhum conflito futuro de agendamento detectado.")
        return

    # Contagem por severidade
    criticos = conflicts_df[conflicts_df["Severidade"] == Config.SEVERITY_CRITICO] if "Severidade" in conflicts_df.columns else pd.DataFrame()
    operacionais = conflicts_df[conflicts_df["Severidade"] == Config.SEVERITY_OPERACIONAL] if "Severidade" in conflicts_df.columns else pd.DataFrame()

    m1, m2, m3 = st.columns(3)
    m1.metric("Total de Conflitos", len(conflicts_df))
    m2.metric("🚨 Conflitos Críticos", len(criticos), help="Férias ou Licença sobreposta com atividade presencial/embarque")
    m3.metric("⚙️ Sobreposições Operacionais", len(operacionais), help="Duas atividades operacionais concorrentes")

    # Botão de download do relatório exclusivo de conflitos
    st.download_button(
        "📥 Baixar Relatório de Conflitos (CSV)",
        conflicts_df.to_csv(index=False).encode("utf-8"),
        "relatorio_conflitos_agenda.csv",
        "text/csv",
        use_container_width=True,
    )

    st.markdown("---")

    df_p80 = conflicts_df[conflicts_df["Projeto"] == "P-80"].reset_index(drop=True)
    df_p82 = conflicts_df[conflicts_df["Projeto"] == "P-82"].reset_index(drop=True)
    df_p83 = conflicts_df[conflicts_df["Projeto"] == "P-83"].reset_index(drop=True)
    df_outros = conflicts_df[~conflicts_df["Projeto"].isin(["P-80", "P-82", "P-83"])].reset_index(drop=True)

    tab_labels = [f"P-80 ({len(df_p80)})", f"P-82 ({len(df_p82)})", f"P-83 ({len(df_p83)})"]
    if not df_outros.empty:
        tab_labels.append(f"Outros ({len(df_outros)})")

    subtabs = st.tabs(tab_labels)
    cols_sub = [
        c for c in [
            "Severidade", "Nome", "Atividade A", "Período A", "Atividade B", "Período B",
            "Período Conflito", "Sobreposição (dias)", "Detalhe_Conflito"
        ] if c in conflicts_df.columns
    ]

    with subtabs[0]:
        if df_p80.empty:
            st.info("Nenhum conflito detectado para P-80.")
        else:
            st.warning(f"⚠️ {len(df_p80)} conflito(s) encontrado(s) para P-80.")
            st.dataframe(df_p80[cols_sub], use_container_width=True, hide_index=True)

    with subtabs[1]:
        if df_p82.empty:
            st.info("Nenhum conflito detectado para P-82.")
        else:
            st.warning(f"⚠️ {len(df_p82)} conflito(s) encontrado(s) para P-82.")
            st.dataframe(df_p82[cols_sub], use_container_width=True, hide_index=True)

    with subtabs[2]:
        if df_p83.empty:
            st.info("Nenhum conflito detectado para P-83.")
        else:
            st.warning(f"⚠️ {len(df_p83)} conflito(s) encontrado(s) para P-83.")
            st.dataframe(df_p83[cols_sub], use_container_width=True, hide_index=True)

    if not df_outros.empty:
        with subtabs[3]:
            st.warning(f"⚠️ {len(df_outros)} conflito(s) em outras bases.")
            st.dataframe(df_outros[cols_sub], use_container_width=True, hide_index=True)


def render_team_tab(display_members: pd.DataFrame, familiar_available: Set[str]) -> None:
    """Renderiza a tabela de visualização de colaboradores com indicador de disponibilidade."""
    status_rows = display_members.copy()
    status_rows["Status no período"] = status_rows["Matrícula"].map(
        lambda m: "Disponível" if str(int(m)) in familiar_available else "Ocupado" if pd.notna(m) else "—"
    )
    show_cols = [
        c for c in [
            "Nome", "Disciplina", "Função", "Projeto", "E-mail", "Matrícula", "Status no período"
        ] if c in status_rows.columns
    ]
    st.dataframe(status_rows[show_cols], use_container_width=True, hide_index=True)


def render_export_tab(
    events_in_window: pd.DataFrame,
    conflicts_df: pd.DataFrame,
    equipe_filtered: pd.DataFrame,
) -> None:
    """Renderiza a aba de dados consolidados com botões de download em formato CSV."""
    st.subheader("📋 Dados Consolidados e Exportação")
    if not events_in_window.empty:
        st.dataframe(events_in_window, use_container_width=True, hide_index=True)
    else:
        st.info("Sem eventos no período selecionado.")

    st.divider()
    c_exp1, c_exp2, c_exp3 = st.columns(3)
    if not events_in_window.empty:
        c_exp1.download_button(
            "📥 Eventos Consolidados (CSV)",
            events_in_window.to_csv(index=False).encode("utf-8"),
            "eventos_consolidados.csv",
            "text/csv",
            use_container_width=True,
        )
    if not conflicts_df.empty:
        c_exp2.download_button(
            "📥 Conflitos Detectados (CSV)",
            conflicts_df.to_csv(index=False).encode("utf-8"),
            "conflitos_agenda.csv",
            "text/csv",
            use_container_width=True,
        )
    equipe_csv = equipe_filtered.to_csv(index=False).encode("utf-8")
    c_exp3.download_button(
        "📥 Equipe Filtrada (CSV)",
        equipe_csv,
        "equipe_filtrada.csv",
        "text/csv",
        use_container_width=True,
    )
