"""
Entrypoint da aplicação Streamlit: Painel de Disponibilidade e Escalas de Plataforma (v0.3).
"""

import warnings
from datetime import datetime
import pandas as pd
import streamlit as st

from escala_p80 import (
    render_aba_escala_projeto,
    render_aba_log_edicoes,
    render_aba_sugestoes,
)
from src.config import Config, VERSION
from src.data_loader import DataLoader
from src.data_processor import DataProcessor
from src.schemas import MissingColumnError
from src.visualizer import Visualizer
from src.ui_components import (
    inject_sidebar_tab_sync,
    render_conflicts_tab,
    render_export_tab,
    render_kpi_cards,
    render_rt_embarques_tab,
    render_sidebar_filters,
    render_team_tab,
)

warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")

# Reexportação para retrocompatibilidade
__all__ = ["App", "Config", "DataLoader", "DataProcessor", "Visualizer", "VERSION", "MissingColumnError"]


class App:
    """Aplicação principal Streamlit que orquestra os dados e renderiza as abas."""

    def __init__(self):
        req_p = st.query_params.get("tab", "")
        initial_state = "collapsed" if req_p and "escala" in str(req_p).lower() else "expanded"
        st.set_page_config(
            layout="wide",
            page_title=f"Gestão de Alocação & Escalas {VERSION}",
            page_icon="📊",
            initial_sidebar_state=initial_state,
        )
        self.hoje = datetime.today().date()

    def run(self):
        st.title(f"🚢 Painel de Disponibilidade e Escalas de Plataforma — {VERSION}")
        st.caption(
            f"Visualização integrada de Estaleiro, Férias, Planejamento Geral e Escalas P-80, P-82 e P-83 ({VERSION})."
        )

        DataLoader.clear_inconsistencias()

        # 1. Carregamento prévio da equipe com tratamento fail-fast amigável
        try:
            equipe_df = DataLoader.load_geral_equipe_data()
        except MissingColumnError as e:
            st.error(f"❌ Erro de Validação de Colunas: {e}")
            return
        except Exception as e:
            st.error(f"❌ Erro inesperado ao carregar dados da equipe: {e}")
            return

        if equipe_df is None:
            st.error("Falha ao carregar a aba Equipe do Planejamento Geral.")
            return

        # 2. Barra lateral com filtros e upload opcional de arquivos
        (
            sel_discs,
            sel_projs,
            sel_funcoes,
            d_inicio,
            d_fim,
            only_available,
            uploaded_files,
        ) = render_sidebar_filters(equipe_df, self.hoje)

        # Se o usuário fez upload de um novo Planejamento Geral, recarrega a equipe a partir dele
        if uploaded_files.get("geral") is not None:
            try:
                equipe_df = DataLoader.load_geral_equipe_data(source=uploaded_files["geral"])
            except MissingColumnError as e:
                st.error(f"❌ Erro no arquivo enviado: {e}")
                return

        equipe_filtered = equipe_df[
            equipe_df["Disciplina"].isin(sel_discs)
            & equipe_df["Projeto"].isin(sel_projs)
            & equipe_df["FunçãoAgrupada"].isin(sel_funcoes)
        ].copy()

        if equipe_filtered.empty:
            st.warning("Nenhum colaborador encontrado com os filtros atuais.")
            return

        # 3. Carregamento dos dados de eventos com suporte a uploads e validação fail-fast
        try:
            _, plan_df = DataLoader.load_estaleiro_data(source=uploaded_files.get("estaleiro"))
            ferias_df = DataLoader.load_ferias_data(source=uploaded_files.get("ferias"))
            geral_df = DataLoader.load_planejamento_geral(source=uploaded_files.get("geral"))
            plataformas_df = DataLoader.load_escalas_plataformas_data(equipe_df)
        except MissingColumnError as e:
            st.error(f"❌ Erro de Validação de Cabeçalhos: {e}")
            return
        except Exception as e:
            st.error(f"❌ Erro ao ler planilhas: {e}")
            return

        # Alerta expansível de inconsistências de datas detectadas
        df_inconsistencias = DataLoader.get_inconsistencias()
        if not df_inconsistencias.empty:
            with st.expander(
                f"⚠️ Inconsistências de Datas Detectadas ({len(df_inconsistencias)} linha(s) ignorada(s))",
                expanded=False,
            ):
                st.warning(
                    "As linhas abaixo possuem valores nulos/inválidos (NaT) ou data de término anterior ao início. "
                    "Elas foram desconsideradas para garantir a integridade dos cálculos do cronograma:"
                )
                st.dataframe(df_inconsistencias, use_container_width=True, hide_index=True)

        # 4. Consolidação de eventos e cálculo vetorizado de conflitos
        combined_df, unique_members = DataProcessor.prepare_combined_data(
            equipe_filtered, [plan_df, ferias_df, geral_df, plataformas_df]
        )

        occupied_ids = DataProcessor.occupied_in_period(combined_df, d_inicio, d_fim)
        familiar_available = DataProcessor.available_in_period(unique_members, combined_df, d_inicio, d_fim)
        events_in_window = DataProcessor.events_in_period(combined_df, d_inicio, d_fim)
        conflicts_full, conflict_keys = DataProcessor.detect_conflicts(combined_df)
        conflicts_futuros = DataProcessor.filter_future_conflicts(conflicts_full, pd.Timestamp(self.hoje))

        display_cols = [
            "Severidade", "Nome", "Projeto", "Atividade A", "Período A", "Atividade B", "Período B",
            "Período Conflito", "Sobreposição (dias)", "Detalhe_Conflito"
        ]
        conflicts_df = (
            conflicts_futuros[[c for c in display_cols if c in conflicts_futuros.columns]].copy()
            if not conflicts_futuros.empty
            else pd.DataFrame(columns=display_cols)
        )

        conflict_details = DataProcessor.build_conflict_details(conflicts_full)

        # 5. Cálculo das métricas de KPIs no Topo (Tarefa 5)
        hoje_ts = pd.Timestamp(self.hoje)
        total_filtrados = len(unique_members)

        # Ativos hoje: colaboradores com evento ativo hoje (exceto férias/folga pura)
        mask_ativos = (
            (combined_df["Início"] <= hoje_ts)
            & (combined_df["Término"] >= hoje_ts)
            & (combined_df["Tipo"].isin(["Estaleiro", "Embarque", "Treinamento", "Workshop", "Visita Técnica"]))
        )
        ativos_hoje = len(set(combined_df.loc[mask_ativos, "Matrícula"].dropna()))

        # Em férias hoje: colaboradores com férias/licença vigentes na data atual
        mask_ferias = (
            (combined_df["Início"] <= hoje_ts)
            & (combined_df["Término"] >= hoje_ts)
            & (combined_df["Tipo"].isin(["Férias", "Licença"]))
        )
        ferias_hoje = len(set(combined_df.loc[mask_ferias, "Matrícula"].dropna()))

        total_conflitos = len(conflicts_df)

        render_kpi_cards(total_filtrados, ativos_hoje, ferias_hoje, total_conflitos)

        # 6. Filtragem para visualização do Gantt
        display_members = unique_members.copy()
        if only_available:
            display_members = display_members[
                display_members["Matrícula"].astype(str).isin(familiar_available)
            ].reset_index(drop=True)

        chart_df = (
            combined_df[combined_df["Matrícula"].isin(set(display_members["Matrícula"]))].copy()
            if not display_members.empty
            else pd.DataFrame()
        )
        chart_df = DataProcessor.events_in_period(chart_df, d_inicio, d_fim) if not chart_df.empty else chart_df

        # 7. Abas da aplicação (Conflitos após as 3 escalas; Equipe e Consolidados visíveis apenas no modo edição)
        is_editor = bool(st.session_state.get("escala_is_editor", False))

        tab_names = [
            "📊 Gráfico de Gantt",
            "Escala P-80",
            "Escala P-82",
            "Escala P-83",
            "⚠️ Conflitos de Agenda",
            "🎫 RT Embarques",
            "Sugestões",
            "Log de Edições",
        ]
        if is_editor:
            tab_names.extend(["Equipe", "📋 Dados Consolidados"])

        req_tab = st.query_params.get("tab") or st.session_state.get("active_main_tab")
        default_tab = None
        if req_tab:
            for t in tab_names:
                if req_tab.lower() in t.lower() or t.lower() in req_tab.lower():
                    default_tab = t
                    break

        if default_tab in tab_names:
            tabs = st.tabs(tab_names, default=default_tab)
        else:
            tabs = st.tabs(tab_names)

        tab_dict = {name: tab for name, tab in zip(tab_names, tabs)}
        inject_sidebar_tab_sync()

        with tab_dict["📊 Gráfico de Gantt"]:
            if display_members.empty:
                st.info("Nenhum colaborador disponível para os critérios selecionados.")
            else:
                y_ticktext = []
                for _, row in display_members.iterrows():
                    nome = str(row["Nome"])
                    mat_str = str(int(row["Matrícula"])) if pd.notna(row["Matrícula"]) else None
                    if mat_str is not None and mat_str in familiar_available:
                        y_ticktext.append(f'<span style="color:#10B981;font-weight:bold">{nome}</span>')
                    else:
                        y_ticktext.append(f'<span style="color:#EF4444;font-weight:bold">{nome}</span>')
                fig = Visualizer.create_gantt_chart(
                    chart_df,
                    display_members,
                    d_inicio,
                    d_fim,
                    conflict_keys,
                    conflict_details,
                    y_ticktext,
                )
                st.plotly_chart(fig, use_container_width=True)
                st.caption(
                    "Verde = disponível no período | Vermelho = ocupado | Barras vermelhas sobrepostas indicam conflitos."
                )

        with tab_dict["Escala P-80"]:
            render_aba_escala_projeto("P-80", equipe_geral_df=equipe_df, cronograma_df=combined_df)

        with tab_dict["Escala P-82"]:
            render_aba_escala_projeto("P-82", equipe_geral_df=equipe_df, cronograma_df=combined_df)

        with tab_dict["Escala P-83"]:
            render_aba_escala_projeto("P-83", equipe_geral_df=equipe_df, cronograma_df=combined_df)

        with tab_dict["⚠️ Conflitos de Agenda"]:
            render_conflicts_tab(conflicts_df, self.hoje)

        with tab_dict["🎫 RT Embarques"]:
            render_rt_embarques_tab(combined_df, self.hoje)

        with tab_dict["Sugestões"]:
            render_aba_sugestoes()

        with tab_dict["Log de Edições"]:
            render_aba_log_edicoes()

        if is_editor:
            if "Equipe" in tab_dict:
                with tab_dict["Equipe"]:
                    render_team_tab(display_members, familiar_available)

            if "📋 Dados Consolidados" in tab_dict:
                with tab_dict["📋 Dados Consolidados"]:
                    render_export_tab(events_in_window, conflicts_df, equipe_filtered)


if __name__ == "__main__":
    App().run()
