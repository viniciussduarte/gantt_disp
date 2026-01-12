"""
Aplicação Streamlit para visualização de alocação de equipe e férias.
Versão Otimizada
"""
import warnings
from datetime import datetime, timedelta, date
from typing import Dict, List, Optional, Tuple, Set

import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st

# Suprimir avisos específicos do Excel
warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")

# --- CONFIGURAÇÃO ---
class Config:
    """Configurações e constantes da aplicação."""
    # Caminhos (Idealmente, mover para st.secrets em produção)
    FILE_PATH_ESTALEIRO = 'Planejamento Estaleiro.xlsx'
    FILE_PATH_FERIAS = 'Férias.xlsx'
    FILE_PATH_GERAL = 'Planejamento Geral.xlsx'

    # Visualização
    EXTENDED_LOOKAHEAD_DAYS = 330
    
    # Cores
    COLOR_TODAY_LINE = "red"
    COLOR_ANALYSIS_PERIOD = "grey"
    COLOR_AVAILABLE = "green"
    COLOR_UNAVAILABLE = "black"
    COLOR_SECTION_LINE = "black"

    # Mapeamento de Atividades
    COLOR_MAP = {
        'Estaleiro': 'blue',
        'Férias': 'red',
        'Folga': 'red',
        'Treinamento': 'orange',
        'Embarque': 'orange',
        'Workshop': 'orange',
        'Visita Técnica': 'orange'
    }

    # Ordem Lógica
    DISCIPLINA_ORDER = ["ELET", "INST", "MEC"]


# --- CARREGAMENTO DE DADOS ---
class DataLoader:
    """Carregamento e normalização de dados."""

    @staticmethod
    def _normalize_dates(df: pd.DataFrame, cols: List[str]) -> pd.DataFrame:
        """Converte colunas para datetime de forma segura e vetorizada."""
        for col in cols:
            if col in df.columns:
                df[col] = pd.to_datetime(df[col], errors='coerce')
        return df

    @staticmethod
    @st.cache_data(ttl=3600) # Cache por 1 hora
    def load_estaleiro_data() -> Tuple[Optional[pd.DataFrame], Optional[pd.DataFrame]]:
        try:
            # Equipe
            equipe_df = pd.read_excel(Config.FILE_PATH_ESTALEIRO, sheet_name='Equipe')
            # Seleção robusta por posição, mas validando nomes
            equipe_df = equipe_df.iloc[:, [0, 1, 3, 4, 5, 7]]
            equipe_df.columns = ['Disciplina', 'Matrícula', 'Função', 'Projeto', 'Experiência', 'Nome']
            equipe_df = equipe_df.dropna(subset=['Experiência']) # Baseado na col 4 original

            # Otimização de memória
            for col in ['Disciplina', 'Função', 'Projeto']:
                equipe_df[col] = equipe_df[col].astype('category')

            # Planejamento
            plan_df = pd.read_excel(
                Config.FILE_PATH_ESTALEIRO, 
                sheet_name='Planejamento IED', 
                skiprows=8, 
                usecols="C:E"
            )
            plan_df.columns = ['Nome', 'Início', 'Término']
            plan_df = plan_df.dropna(subset=['Nome'])
            plan_df = DataLoader._normalize_dates(plan_df, ['Início', 'Término'])

            # Merge Otimizado
            plan_df = plan_df.merge(
                equipe_df[['Nome', 'Matrícula', 'Disciplina', 'Função', 'Projeto']],
                on='Nome',
                how='left'
            )
            plan_df['Tipo'] = 'Estaleiro'

            return equipe_df, plan_df

        except Exception as e:
            st.error(f"Erro ao carregar Estaleiro: {e}")
            return None, None

    @staticmethod
    @st.cache_data(ttl=3600)
    def load_ferias_data() -> Optional[pd.DataFrame]:
        try:
            df = pd.read_excel(Config.FILE_PATH_FERIAS, skiprows=1, header=None)
            
            # 1. Seleção das colunas (Matrícula + 3 parcelas de Início/Término)
            cols_idx = [0, 8, 9, 11, 12, 14, 15]
            col_names = [
                "Matrícula", 
                "Início_1", "Término_1", 
                "Início_2", "Término_2", 
                "Início_3", "Término_3"
            ]
            df = df.iloc[:, cols_idx].copy()
            df.columns = col_names

            # --- SOLUÇÃO PARA O ERRO DE ID ÚNICO ---
            # Criamos um ID de linha único para que o Pandas saiba diferenciar 
            # registros diferentes da mesma matrícula durante o "melt"
            df['row_id'] = range(len(df))

            # 2. Reshape (Wide to Long) usando 'row_id' e 'Matrícula' como identificadores
            df_long = pd.wide_to_long(
                df, 
                stubnames=["Início", "Término"], 
                i=["row_id", "Matrícula"], # O par (row_id, Matrícula) agora é único
                j="Parcela", 
                sep="_", 
                suffix=r'\d+'
            ).reset_index()

            # 3. Limpeza e Normalização
            df_long = df_long.dropna(subset=["Início"])
            df_long['Tipo'] = "Férias"
            
            # Converte datas de forma segura
            df_long = DataLoader._normalize_dates(df_long, ['Início', 'Término'])
            
            # Removemos o row_id pois ele não é mais necessário após o processamento
            df_long.drop(columns=['row_id', 'Parcela'], inplace=True)
            
            return df_long

        except Exception as e:
            st.error(f"Erro ao carregar Férias: {e}")
            return None
        
    @staticmethod
    @st.cache_data(ttl=3600)
    def load_planejamento_geral() -> Optional[pd.DataFrame]:
        try:
            df = pd.read_excel(
                Config.FILE_PATH_GERAL,
                usecols=["Nome", "Matrícula", "Início", "Término", "Atividade", "Detalhamento"]
            )
            df = df.rename(columns={'Atividade': 'Tipo'})
            df = DataLoader._normalize_dates(df, ['Início', 'Término'])
            return df
        except Exception as e:
            st.error(f"Erro ao carregar Planejamento Geral: {e}")
            return None


# --- PROCESSAMENTO ---
class DataProcessor:
    """Lógica de negócios e manipulação de dados."""

    @staticmethod
    def prepare_combined_data(equipe_df: pd.DataFrame, 
                              dfs_eventos: List[pd.DataFrame]) -> Tuple[pd.DataFrame, pd.DataFrame]:
        """Combina todas as fontes de dados em um único DataFrame normalizado."""
        
        valid_dfs = []
        matriculas_validas = set(equipe_df['Matrícula'])

        for df in dfs_eventos:
            if df is not None and not df.empty:
                # Filtrar apenas matrículas que existem na equipe atual
                df_filtered = df[df['Matrícula'].isin(matriculas_validas)].copy()
                
                # Garantir colunas essenciais
                cols_to_keep = ['Matrícula', 'Início', 'Término', 'Tipo']
                if 'Nome' in df.columns: cols_to_keep.append('Nome')
                if 'Detalhamento' in df.columns: cols_to_keep.append('Detalhamento')
                
                valid_dfs.append(df_filtered[cols_to_keep])

        if not valid_dfs:
            combined = pd.DataFrame(columns=['Matrícula', 'Nome', 'Início', 'Término', 'Disciplina', 'Tipo'])
        else:
            combined = pd.concat(valid_dfs, ignore_index=True)

        # Enriquecer com dados da equipe (Merge é mais rápido que iterrows)
        combined = combined.merge(
            equipe_df[['Matrícula', 'Nome', 'Disciplina', 'Função', 'Projeto']],
            on='Matrícula',
            how='left',
            suffixes=('', '_eq')
        )
        
        # Preencher Nome faltante se necessário
        if 'Nome_eq' in combined.columns:
            combined['Nome'] = combined['Nome'].fillna(combined['Nome_eq'])
            combined.drop(columns=['Nome_eq'], inplace=True)

        # Ordenação para o Gráfico
        disciplina_type = pd.CategoricalDtype(Config.DISCIPLINA_ORDER, ordered=True)
        
        # DataFrame de Membros Únicos Ordenados
        unique_members = equipe_df.copy()
        unique_members['Disciplina'] = unique_members['Disciplina'].astype(disciplina_type)
        unique_members = unique_members.sort_values(
            by=['Disciplina', 'Função', 'Projeto', 'Nome']
        ).reset_index(drop=True)

        # Ajustar combinado
        if not combined.empty:
            combined['Disciplina'] = combined['Disciplina'].astype(disciplina_type)

        return combined, unique_members

    @staticmethod
    def get_available_members(equipe_df: pd.DataFrame, combined_df: pd.DataFrame, 
                             start_date: pd.Timestamp, end_date: pd.Timestamp) -> pd.DataFrame:
        """
        Retorna apenas membros sem alocação no período.
        Usa lógica de conjuntos para performance O(1) na verificação.
        """
        if combined_df.empty:
            return equipe_df

        # Filtrar eventos que colidem com a janela de análise
        mask_periodo = (
            (combined_df['Início'] <= end_date) & 
            (combined_df['Término'] >= start_date)
        )
        ocupados_ids = set(combined_df.loc[mask_periodo, 'Matrícula'].unique())
        
        # Retorna apenas quem NÃO está no set de ocupados
        return equipe_df[~equipe_df['Matrícula'].isin(ocupados_ids)].copy()

    @staticmethod
    def detect_conflicts_vectorized(combined_df: pd.DataFrame) -> pd.DataFrame:
        """
        Detecta conflitos usando operações vetorizadas (shift) em vez de loops.
        Retorna um DataFrame com os conflitos.
        """
        if combined_df.empty:
            return pd.DataFrame()

        df = combined_df.sort_values(by=['Matrícula', 'Início'])
        
        # Criar colunas deslocadas para comparar linha atual com a próxima
        df['Next_Inicio'] = df.groupby('Matrícula')['Início'].shift(-1)
        df['Next_Tipo'] = df.groupby('Matrícula')['Tipo'].shift(-1)
        
        # Lógica de conflito: Término Atual > Próximo Início (dentro da mesma matrícula)
        # Nota: Ajustar > ou >= dependendo se término no dia X e início no dia X é conflito.
        # Assumindo que sim para segurança.
        conflict_mask = (df['Término'] > df['Next_Inicio']) & (df['Next_Inicio'].notna())
        
        conflicts = df[conflict_mask].copy()
        
        if conflicts.empty:
            return pd.DataFrame()

        # Formatar saída
        saida = []
        for _, row in conflicts.iterrows():
            saida.append({
                "Nome": row['Nome'],
                "Conflito": f"{row['Tipo']} ({row['Término'].strftime('%d/%m')}) x {row['Next_Tipo']} ({row['Next_Inicio'].strftime('%d/%m')})"
            })
            
        return pd.DataFrame(saida)


# --- VISUALIZAÇÃO ---
class Visualizer:
    @staticmethod
    def create_gantt_chart(combined_df: pd.DataFrame, unique_members: pd.DataFrame,
                          start_date: pd.Timestamp, end_date: pd.Timestamp, 
                          occupied_ids: Set) -> go.Figure:
        
        y_order = unique_members['Nome'].tolist()
        
        # Criar labels coloridos HTML
        y_labels = []
        for nome, mat in zip(unique_members['Nome'], unique_members['Matrícula']):
            color = Config.COLOR_UNAVAILABLE if mat in occupied_ids else Config.COLOR_AVAILABLE
            weight = "bold" if color == Config.COLOR_AVAILABLE else "normal"
            y_labels.append(f'<span style="color:{color}; font-weight:{weight}">{nome}</span>')

        # Gráfico Base
        fig = px.timeline(
            combined_df,
            x_start="Início", x_end="Término", y="Nome",
            color="Tipo",
            color_discrete_map=Config.COLOR_MAP,
            category_orders={"Nome": y_order},
            hover_data=["Disciplina", "Projeto", "Detalhamento"]
        )

        # Estilização
        fig.update_traces(marker=dict(line=dict(width=1, color='black')), selector=dict(type='bar'))
        
        # Layout
        fig.update_layout(
            height=max(600, len(y_order) * 30), # Altura dinâmica
            xaxis_range=[start_date - timedelta(days=2), end_date + timedelta(days=2)],
            yaxis=dict(
                tickmode='array', tickvals=y_order, ticktext=y_labels,
                gridcolor='lightgrey'
            ),
            xaxis=dict(gridcolor='lightgrey', title="Data"),
            plot_bgcolor='white',
            title="Cronograma de Alocação",
            legend_title="Atividade"
        )

        # Linha "Hoje"
        hoje = datetime.now()
        fig.add_vline(x=hoje.timestamp() * 1000, line_width=2, line_dash="dash", line_color="red", annotation_text="Hoje")
        
        # Divisores de Disciplina
        y_pos = 0
        for disc in reversed(Config.DISCIPLINA_ORDER):
            count = len(unique_members[unique_members['Disciplina'] == disc])
            if count > 0:
                y_pos += count
                fig.add_hline(y=y_pos - 0.5, line_dash="dot", line_color="black")
                fig.add_annotation(x=1, y=y_pos - (count/2) - 0.5, text=f"<b>{disc}</b>", 
                                 xref="paper", yref="y", xanchor="right", showarrow=False)

        return fig


# --- APLICAÇÃO PRINCIPAL ---
class App:
    def __init__(self):
        st.set_page_config(layout="wide", page_title="Gestão de Alocação")
        self.hoje = datetime.today().date()

    def run(self):
        st.title("📊 Painel de Alocação de Equipe")

        # 1. Sidebar e Filtros
        with st.sidebar:
            st.header("Filtros")
            
            # Carregar dados
            equipe_df, plan_df = DataLoader.load_estaleiro_data()
            ferias_df = DataLoader.load_ferias_data()
            geral_df = DataLoader.load_planejamento_geral()

            if equipe_df is None:
                st.error("Falha ao carregar arquivo principal (Estaleiro).")
                return

            # Filtros Dinâmicos
            all_discs = sorted(equipe_df['Disciplina'].unique())
            sel_discs = st.multiselect("Disciplina", all_discs, default=all_discs)
            
            all_projs = sorted(equipe_df['Projeto'].unique())
            sel_projs = st.multiselect("Projeto", all_projs, default=all_projs)

            # Datas
            col1, col2 = st.columns(2)
            d_inicio = pd.Timestamp(col1.date_input("Início", self.hoje - timedelta(days=7)))
            d_fim = pd.Timestamp(col2.date_input("Fim", self.hoje + timedelta(days=90)))

            only_available = st.checkbox("Apenas Disponíveis", help="Mostra quem não tem nada agendado no período")

        # 2. Processamento
        # Filtragem inicial da equipe
        equipe_filtered = equipe_df[
            (equipe_df['Disciplina'].isin(sel_discs)) & 
            (equipe_df['Projeto'].isin(sel_projs))
        ].copy()

        if equipe_filtered.empty:
            st.warning("Nenhum colaborador encontrado com os filtros atuais.")
            return

        # Combinar eventos (apenas para a equipe filtrada para economizar processamento)
        combined_df, unique_members = DataProcessor.prepare_combined_data(
            equipe_filtered, [plan_df, ferias_df, geral_df]
        )

        # Lógica de Disponibilidade (Set-based, muito rápida)
        # Identificar IDs ocupados no período selecionado
        mask_ocupados = (combined_df['Início'] <= d_fim) & (combined_df['Término'] >= d_inicio)
        occupied_ids = set(combined_df.loc[mask_ocupados, 'Matrícula'].unique())

        if only_available:
            # Filtra unique_members para manter apenas quem NÃO está no set occupied_ids
            unique_members = unique_members[~unique_members['Matrícula'].isin(occupied_ids)]
            # Refiltra o combined para o gráfico não mostrar barras de quem foi removido
            combined_df = combined_df[combined_df['Matrícula'].isin(unique_members['Matrícula'])]

        if unique_members.empty:
            st.info("Nenhum colaborador disponível para os critérios selecionados.")
            return

        # 3. Visualização
        tab_grafico, tab_conflitos = st.tabs(["Cronograma", "Relatório de Conflitos"])

        with tab_grafico:
            # Filtrar dados para o gráfico (apenas o necessário)
            chart_df = combined_df[combined_df['Matrícula'].isin(unique_members['Matrícula'])].copy()
            
            # Ordenação do gráfico baseada na lista de membros únicos
            chart_df['Nome'] = pd.Categorical(
                chart_df['Nome'], 
                categories=unique_members['Nome'], 
                ordered=True
            )
            chart_df = chart_df.sort_values('Nome')

            fig = Visualizer.create_gantt_chart(
                chart_df, unique_members, d_inicio, d_fim, occupied_ids
            )
            st.plotly_chart(fig, use_container_width=True)

        with tab_conflitos:
            conflicts_df = DataProcessor.detect_conflicts_vectorized(combined_df)
            if conflicts_df.empty:
                st.success("✅ Nenhum conflito de agendamento detectado.")
            else:
                st.warning(f"⚠️ {len(conflicts_df)} Conflitos Encontrados")
                st.dataframe(conflicts_df, use_container_width=True, hide_index=True)


if __name__ == "__main__":
    App().run()