"""
Aplicação Streamlit para visualização de alocação de equipe e férias.
"""
import warnings
from datetime import datetime, timedelta
from typing import Dict, List, Optional, Tuple, Union

import pandas as pd
import plotly.express as px
import streamlit as st

import plotly.graph_objects as go 
from plotly.graph_objects import Figure as PlotlyFigure


# Suprimir avisos específicos
warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")

# Configurações e constantes
class Config:
    """Configurações da aplicação."""
    # Caminhos dos arquivos
    FILE_PATH_ESTALEIRO = 'Planejamento Estaleiro.xlsx'
    FILE_PATH_FERIAS = 'Férias.xlsx'
    FILE_PATH_GERAL = 'Planejamento Geral.xlsx'
    
    # Configurações de visualização
    DEFAULT_LOOKBACK_DAYS = 30
    DEFAULT_LOOKAHEAD_DAYS = 90
    EXTENDED_LOOKAHEAD_DAYS = 330
    
    # Cores
    COLOR_TODAY_LINE = "red"
    COLOR_ANALYSIS_PERIOD = "grey"
    COLOR_AVAILABLE = "green"
    COLOR_UNAVAILABLE = "black"
    COLOR_SECTION_LINE = "black"
    
    # Mapeamento de cores para os tipos de atividades
    COLOR_MAP = {
        'Estaleiro': 'blue',
        'Férias': 'red',
        'Folga': 'red',
        'Treinamento':'orange',
        'Embarque':'orange',
        'Workshop':'orange',
        'Visita Técnica':'orange'
    }

    # Ordem das disciplinas (adicionado aqui para fácil acesso)
    DISCIPLINA_ORDER = ["ELET", "INST", "MEC"]


class DataLoader:
    """Classe responsável por carregar e processar os dados."""
    
    @staticmethod
    @st.cache_data
    def load_estaleiro_data() -> Tuple[Optional[pd.DataFrame], Optional[pd.DataFrame]]:
        """
        Carrega dados do arquivo de estaleiro.
        
        Returns:
            Tuple contendo DataFrame de equipe e DataFrame de planejamento.
        """
        try:
            # Leitura da aba "Equipe"
            equipe_df = pd.read_excel(
                Config.FILE_PATH_ESTALEIRO, 
                sheet_name='Equipe'
            )
            equipe_df = equipe_df.iloc[:, [0, 1, 3, 4, 5, 7]].dropna(subset=[equipe_df.columns[4]])
            equipe_df.columns = ['Disciplina', 'Matrícula', 'Função', 'Projeto', 'Experiência', 'Nome']
            
            # Leitura da aba "Planejamento IED"
            planejamento_df = pd.read_excel(
                Config.FILE_PATH_ESTALEIRO, 
                sheet_name='Planejamento IED', 
                skiprows=8, 
                usecols="C:E"
            )
            planejamento_df = planejamento_df.dropna(subset=[planejamento_df.columns[0]])
            planejamento_df.columns = ['Nome', 'Início', 'Término']
            
            # Converter datas para formato datetime
            planejamento_df['Início'] = pd.to_datetime(planejamento_df['Início'])
            planejamento_df['Término'] = pd.to_datetime(planejamento_df['Término'])
            
            # Merge com "Equipe" para trazer a matrícula
            planejamento_df = pd.merge(
                planejamento_df, 
                equipe_df[['Nome', 'Matrícula', 'Disciplina', 'Função', 'Projeto']],
                on='Nome', 
                how='left'
            )
            planejamento_df['Tipo'] = 'Estaleiro'
            
            return equipe_df, planejamento_df
            
        except FileNotFoundError:
            st.error(f"Arquivo {Config.FILE_PATH_ESTALEIRO} não encontrado. Por favor, verifique o caminho do arquivo.")
            return None, None
        except Exception as e:
            st.error(f"Erro ao carregar dados do estaleiro: {str(e)}")
            return None, None
    
    @staticmethod
    @st.cache_data
    def load_ferias_data() -> Optional[pd.DataFrame]:
        """
        Carrega dados do arquivo de férias.
        
        Returns:
            DataFrame com dados de férias processados.
        """
        try:
            df_ferias = pd.read_excel(Config.FILE_PATH_FERIAS, skiprows=1, header=None)
            
            # Definir colunas manualmente
            df_ferias.columns = [
                "Matrícula", "Nome do Empregado", "Período Aquisitivo", "Prazo Mínimo",
                "Prazo Máximo", "13º Salário", "Abono", "Período Único",
                "Primeira Parcela", "Termino Primeira Parcela", "Nº de Dias 1",
                "Segunda Parcela", "Termino Segunda Parcela", "Nº de Dias 2",
                "Terceira Parcela", "Termino Terceira Parcela", "Nº de Dias 3",
                "Situação"
            ]
            
            # Selecionar colunas relevantes
            colunas_indices = [0, 8, 9, 11, 12, 14, 15]
            df_selecionado = df_ferias.iloc[:, colunas_indices]
            df_selecionado.columns = [
                "Matrícula",
                "Primeira Parcela", "Termino Primeira Parcela",
                "Segunda Parcela", "Termino Segunda Parcela",
                "Terceira Parcela", "Termino Terceira Parcela"
            ]
            
            # Processar cada parcela de férias
            parcelas = []
            
            for parcela, termino in [
                ("Primeira Parcela", "Termino Primeira Parcela"),
                ("Segunda Parcela", "Termino Segunda Parcela"),
                ("Terceira Parcela", "Termino Terceira Parcela")
            ]:
                df_parcela = df_selecionado[["Matrícula", parcela, termino]].dropna(subset=[parcela])
                df_parcela = df_parcela.rename(columns={parcela: "Início", termino: "Término"})
                df_parcela["Tipo"] = "Férias"
                parcelas.append(df_parcela)
            
            # Concatenar as parcelas de férias
            df_ferias_final = pd.concat(parcelas, ignore_index=True)
            df_ferias_final['Início'] = pd.to_datetime(df_ferias_final['Início'])
            df_ferias_final['Término'] = pd.to_datetime(df_ferias_final['Término'])
            
            return df_ferias_final
            
        except FileNotFoundError:
            st.error(f"Arquivo {Config.FILE_PATH_FERIAS} não encontrado. Por favor, verifique o caminho do arquivo.")
            return None
        except Exception as e:
            st.error(f"Erro ao carregar dados de férias: {str(e)}")
            return None
    
    @staticmethod
    @st.cache_data
    def load_planejamento_geral() -> Optional[pd.DataFrame]:
        """
        Carrega dados do arquivo de planejamento geral.
        
        Returns:
            DataFrame com dados de planejamento geral.
        """
        try:
            planejamento_geral_df = pd.read_excel(
                Config.FILE_PATH_GERAL, 
                usecols=["Nome", "Matrícula", "Início", "Término", "DIAS", "Atividade", "Detalhamento"]
            )
            
            planejamento_geral_df = planejamento_geral_df.rename(columns={'Atividade': 'Tipo'})
            planejamento_geral_df[['Início', 'Término']] = planejamento_geral_df[['Início', 'Término']].apply(pd.to_datetime)
            
            return planejamento_geral_df
            
        except FileNotFoundError:
            st.error(f"Arquivo {Config.FILE_PATH_GERAL} não encontrado. Por favor, verifique o caminho do arquivo.")
            return None
        except Exception as e:
            st.error(f"Erro ao carregar dados de planejamento geral: {str(e)}")
            return None


class DataProcessor:
    """Classe para processamento e análise de dados."""
    
    @staticmethod
    def check_conflict(row: pd.Series, combined_df: pd.DataFrame, 
                       start_date: pd.Timestamp, end_date: pd.Timestamp) -> bool:
        """
        Verifica se há conflitos de agenda para uma pessoa em um período específico.
        
        Args:
            row: Linha do DataFrame contendo dados da pessoa.
            combined_df: DataFrame combinado com todos os eventos.
            start_date: Data de início da análise.
            end_date: Data de fim da análise.
            
        Returns:
            True se existir conflito, False caso contrário.
        """
        periodos = combined_df[
            (combined_df['Matrícula'] == row['Matrícula']) &
            (combined_df['Início'] <= end_date) &
            (combined_df['Término'] >= start_date)
        ]
        return not periodos.empty
    
    @staticmethod
    def detect_conflicts(combined_df: pd.DataFrame) -> str:
        """
        Detecta conflitos de agenda para todas as pessoas no DataFrame combinado.
        
        Args:
            combined_df: DataFrame combinado com todos os eventos.
            
        Returns:
            Texto formatado com os conflitos detectados.
        """
        conflitos_texto = []
        
        for matricula in combined_df['Matrícula'].unique():
            eventos = combined_df[
                (combined_df['Matrícula'] == matricula)
            ].sort_values(by='Início')
            
            if len(eventos) <= 1:
                continue
            
            nome = eventos['Nome'].iloc[0] if 'Nome' in eventos.columns and not eventos['Nome'].empty else f"Matrícula: {matricula}"
            
            for i in range(len(eventos) - 1):
                evento_atual = eventos.iloc[[i]]
                proximo_evento = eventos.iloc[[i+1]]
                
                termino_atual = evento_atual['Término'].iloc[0]
                inicio_proximo = proximo_evento['Início'].iloc[0]
                
                # Verifica se há um conflito de fato (se as datas se sobrepõem)
                if termino_atual >= inicio_proximo and termino_atual.date() > inicio_proximo.date():
                    linha = (
                        f"{nome} - {evento_atual['Tipo'].iloc[0]} "
                        f"({evento_atual['Início'].iloc[0].strftime('%d/%m/%Y')} a {termino_atual.strftime('%d/%m/%Y')}) / "
                        f"{proximo_evento['Tipo'].iloc[0]} "
                        f"({inicio_proximo.strftime('%d/%m/%Y')} a {proximo_evento['Término'].iloc[0].strftime('%d/%m/%Y')})"
                    )
                    conflitos_texto.append(linha)
        
        if conflitos_texto:
            return "CONFLITOS DETECTADOS:\n" + "\n".join(conflitos_texto)
        else:
            return "Nenhum conflito detectado."
    
    @staticmethod
    def prepare_combined_data(equipe_df: pd.DataFrame, 
                              planejamento_df: Optional[pd.DataFrame],
                              ferias_df: Optional[pd.DataFrame],
                              planejamento_geral_df: Optional[pd.DataFrame]) -> Tuple[pd.DataFrame, pd.DataFrame]:
        """
        Prepara os dados combinados para visualização.
        
        Args:
            equipe_df: DataFrame com dados da equipe.
            planejamento_df: DataFrame com dados de planejamento.
            ferias_df: DataFrame com dados de férias.
            planejamento_geral_df: DataFrame com dados de planejamento geral.
            
        Returns:
            Tuple contendo DataFrame combinado e DataFrame de membros únicos.
        """
        dataframes_to_combine = []
        
        if planejamento_df is not None:
            # Filtra o planejamento para incluir apenas membros da equipe filtrada
            planejamento_df_filtrado = planejamento_df[
                planejamento_df['Matrícula'].isin(equipe_df['Matrícula'])
            ].copy()
            if not planejamento_df_filtrado.empty:
                dataframes_to_combine.append(
                    planejamento_df_filtrado[['Matrícula', 'Nome', 'Início', 'Término', 'Disciplina', 'Função', 'Projeto', 'Tipo']]
                )
        
        if ferias_df is not None:
            ferias_df_filtrado = ferias_df[
                ferias_df['Matrícula'].isin(equipe_df['Matrícula'])
            ].copy()
            
            # ATENÇÃO: Corrigindo aqui para pegar o nome da aba Equipe
            if not ferias_df_filtrado.empty:
                ferias_com_nome = pd.merge(
                    ferias_df_filtrado, 
                    equipe_df[['Matrícula', 'Nome']],
                    on='Matrícula', 
                    how='left'
                )
                dataframes_to_combine.append(ferias_com_nome[['Matrícula', 'Nome', 'Início', 'Término', 'Tipo']])
        
        if planejamento_geral_df is not None:
            planejamento_geral_df_filtrado = planejamento_geral_df[
                planejamento_geral_df['Matrícula'].isin(equipe_df['Matrícula'])
            ].copy()
            if not planejamento_geral_df_filtrado.empty:
                dataframes_to_combine.append(
                    planejamento_geral_df_filtrado[['Matrícula', 'Nome', 'Início', 'Término', 'Tipo', 'Detalhamento']]
                )
        
        if not dataframes_to_combine:
            combined_df = pd.DataFrame(
                columns=['Matrícula', 'Nome', 'Início', 'Término', 'Disciplina', 'Função', 'Projeto', 'Tipo']
            )
        else:
            combined_df = pd.concat(dataframes_to_combine, ignore_index=True)
        
        # Criar mapeamento de matrículas para informações da equipe
        matricula_mapping = {row['Matrícula']: row for _, row in equipe_df.iterrows()}
        
        # Garantir que todas as colunas necessárias existem
        for col in ['Disciplina', 'Função', 'Projeto']:
            if col not in combined_df.columns:
                combined_df[col] = None
        
        # Preencher informações faltantes
        for idx, row in combined_df.iterrows():
            if pd.isna(row['Disciplina']) or pd.isna(row['Função']) or pd.isna(row['Projeto']):
                if row['Matrícula'] in matricula_mapping:
                    emp_data = matricula_mapping[row['Matrícula']]
                    combined_df.at[idx, 'Disciplina'] = emp_data['Disciplina']
                    combined_df.at[idx, 'Função'] = emp_data['Função']
                    combined_df.at[idx, 'Projeto'] = emp_data['Projeto']
        
        # Preparar membros únicos
        disciplinas_ordenadas = pd.CategoricalDtype(Config.DISCIPLINA_ORDER, ordered=True)
        unique_members = equipe_df.copy()
        unique_members['Disciplina'] = unique_members['Disciplina'].astype(disciplinas_ordenadas)
        
        unique_members = unique_members.sort_values(
            by=['Disciplina', 'Função', 'Projeto', 'Nome'], 
            na_position='last'
        ).reset_index(drop=True)

        # Reordenar combined_df com a mesma lógica
        if 'Disciplina' in combined_df.columns:
            combined_df['Disciplina'] = combined_df['Disciplina'].astype(disciplinas_ordenadas)
        
        return combined_df, unique_members


class Visualizer:
    """Classe para criação de visualizações."""
    
    @staticmethod
    def create_gantt_chart(combined_df_filtered: pd.DataFrame, 
                           unique_members: pd.DataFrame, 
                           start_date: pd.Timestamp, 
                           end_date: pd.Timestamp,
                           hoje: datetime.date) -> go.Figure:
        """
        Cria um gráfico de Gantt para visualização de alocação.
        
        Args:
            combined_df_filtered: DataFrame filtrado e combinado.
            unique_members: DataFrame com membros únicos.
            start_date: Data de início da análise.
            end_date: Data de fim da análise.
            hoje: Data atual.
            
        Returns:
            Figura Plotly com o gráfico de Gantt.
        """
        # Crie a lista ordenada de rótulos do eixo Y com nome e projeto
        y_order = unique_members['Nome'].tolist()

        # Crie um dicionário com a cor para cada pessoa
        cor_nomes_dict = {}
        for _, row in unique_members.iterrows():
            disponivel = not DataProcessor.check_conflict(row, combined_df_filtered, start_date, end_date)
            cor = Config.COLOR_AVAILABLE if disponivel else Config.COLOR_UNAVAILABLE
            cor_nomes_dict[row['Nome']] = cor
        
        # Crie os rótulos do eixo Y usando a formatação HTML com as cores
        y_ticktext_colored = [
            f'<span style="color:{cor_nomes_dict.get(nome, "black")}">{nome}</span>' 
            for nome in y_order
        ]
        
        # Crie o gráfico de Gantt com Plotly Express
        fig = px.timeline(
            combined_df_filtered,
            x_start="Início",
            x_end="Término",
            y="Nome",
            color="Tipo",
            color_discrete_map=Config.COLOR_MAP,
            hover_data={
                "Nome": True,
                "Detalhamento": True,
                "Início": "|%d/%m/%Y",
                "Término": "|%d/%m/%Y",
                "Tipo": True,
                "Disciplina": True,
                "Função": True,
                "Projeto": True
            },
            category_orders={"Nome": y_order}
        )
        
        # Adicionar borda preta às barras
        fig.update_traces(
            marker=dict(
                line=dict(
                    width=1,
                    color='black'
                )
            ),
            selector=dict(type='bar')
        )
        
        # Ajustes para melhor visibilidade
        fig.update_layout(
            template='plotly_white',
            plot_bgcolor='white',
            paper_bgcolor='white',
            font_color='black',
            yaxis=dict(
                tickmode='array',
                tickvals=y_order,
                ticktext=y_ticktext_colored,
                gridcolor='lightgrey',
                tickfont=dict(color='black')
            ),
            xaxis=dict(
                gridcolor='lightgrey',
                tickfont=dict(color='black')
            ),
            title_text='Alocação da Equipe e Férias',
            xaxis_title="Data",
            yaxis_title="Nome",
            height=800,
            xaxis_range=[
                (hoje - timedelta(days=30)), 
                (hoje + timedelta(days=Config.EXTENDED_LOOKAHEAD_DAYS))
            ],
            # Garantir que o título da legenda seja preto
            legend_title=dict(
                text="Atividade",
                font=dict(color='black', size=12)
            ),
            legend=dict(
                font=dict(color='black'),
                title_font=dict(color='black')
            )
        )
        
        # Adicionar linha vertical para hoje
        hoje_datetime = datetime.combine(hoje, datetime.min.time())
        fig.add_vline(
            x=hoje_datetime.timestamp() * 1000, 
            line_width=2, 
            line_dash="dash", 
            line_color=Config.COLOR_TODAY_LINE, 
            annotation_text="Hoje", 
            annotation_position="top left",
            annotation=dict(
                font=dict(color=Config.COLOR_TODAY_LINE),
                yref="paper",
                y=1.025, # Posiciona a anotação acima do gráfico
                showarrow=False
            )
        )

        # Adicionar retângulo para o período de análise
        fig.add_vrect(
            x0=start_date, 
            x1=end_date,
            fillcolor=Config.COLOR_ANALYSIS_PERIOD, 
            opacity=0.2, 
            line_width=0,
        )

        # Adicionar bordas para as disciplinas e suas legendas na ordem correta
        y_posicao_atual = 0.0
        for disc in reversed(Config.DISCIPLINA_ORDER): # Mudando a ordem para mostrar ELET no topo
            members_da_disciplina = unique_members[unique_members['Disciplina'] == disc]
            if not members_da_disciplina.empty:
                num_membros = len(members_da_disciplina)
                ultima_posicao = y_posicao_atual + num_membros
                
                # Adiciona a linha divisória
                fig.add_hline(
                    y=ultima_posicao - 0.5, 
                    line_width=2, 
                    line_dash="dash", 
                    line_color=Config.COLOR_SECTION_LINE
                )
                
                # Adiciona a anotação (legenda) da disciplina
                y_legenda = y_posicao_atual + (num_membros / 2)
                fig.add_annotation(
                    # Coordenadas relativas ao "papel" do gráfico
                    xref="paper",
                    yref="y", # Mantém a referência do eixo Y para a posição vertical
                    x=1,      # Posição X no extremo direito (1.0 = borda direita)
                    y=y_legenda,
                    text=f"<b>{disc}</b>",
                    showarrow=False,
                    xanchor="right", # Alinha a borda direita da anotação com a posição x=1
                    yanchor="middle",
                    font=dict(size=14, color=Config.COLOR_SECTION_LINE)
                )

                # Atualiza a posição para o próximo grupo de membros
                y_posicao_atual = ultima_posicao
        
        return fig

class App:
    """Classe principal da aplicação Streamlit."""
    
    def __init__(self):
        """Inicializa a aplicação."""
        self.hoje = datetime.today().date()
        
    def setup_page(self):
        """Configura a página Streamlit."""
        st.set_page_config(layout="wide", page_title="Alocação de Equipe e Férias")
        st.title("Alocação de Equipe e Férias")
        
    def run(self):
        """Executa a aplicação principal."""
        self.setup_page()
        
        # Configuração da barra lateral
        st.sidebar.header("Configurações de Análise")
        
        # Carregar dados
        equipe_df, planejamento_df = DataLoader.load_estaleiro_data()
        df_ferias_final = DataLoader.load_ferias_data()
        planejamento_geral_df = DataLoader.load_planejamento_geral()
        
        if equipe_df is not None:
            # Filtros de disciplina e projeto (plataforma)
            disciplinas_unicas = sorted(equipe_df['Disciplina'].dropna().unique())
            disciplinas_selecionadas = st.sidebar.multiselect(
                "Filtrar por Disciplina",
                options=disciplinas_unicas,
                default=disciplinas_unicas
            )

            projetos_unicos = sorted(equipe_df['Projeto'].dropna().unique())
            projetos_selecionados = st.sidebar.multiselect(
                "Filtrar por Plataforma (Projeto)",
                options=projetos_unicos,
                default=projetos_unicos
            )
            
            data_inicio_analise = st.sidebar.date_input(
                "Data Início Análise", 
                value=self.hoje - timedelta(days=Config.DEFAULT_LOOKBACK_DAYS)
            )
            data_fim_analise = st.sidebar.date_input(
                "Data Fim Análise", 
                value=self.hoje + timedelta(days=Config.DEFAULT_LOOKAHEAD_DAYS)
            )

            # Converter para timestamp
            data_inicio_analise = pd.Timestamp(data_inicio_analise)
            data_fim_analise = pd.Timestamp(data_fim_analise)
            
            # Aplicar os filtros ao DataFrame da equipe
            equipe_df_filtrada = equipe_df[
                (equipe_df['Disciplina'].isin(disciplinas_selecionadas)) &
                (equipe_df['Projeto'].isin(projetos_selecionados))
            ].copy()

            # Verificar se o DataFrame filtrado não está vazio antes de continuar
            if equipe_df_filtrada.empty:
                st.warning("Nenhum membro da equipe corresponde aos filtros selecionados. Por favor, ajuste suas seleções.")
                return 
            
            # Preparar dados combinados
            combined_df, unique_members = DataProcessor.prepare_combined_data(
                equipe_df_filtrada, planejamento_df, df_ferias_final, planejamento_geral_df
            )
            
            # Filtrar e ordenar dados
            matriculas_equipe = unique_members['Matrícula'].unique()
            combined_df_filtered = combined_df[
                (combined_df['Matrícula'].isin(matriculas_equipe))
            ].copy()
            
            # Criar nome completo e ordenar
            unique_members_list = unique_members['Nome'].tolist()
            
            combined_df_filtered['Nome'] = pd.Categorical(
                combined_df_filtered['Nome'], 
                categories=unique_members_list, 
                ordered=True
            )
            combined_df_filtered = combined_df_filtered.sort_values('Nome')

            # Criar visualização
            fig = Visualizer.create_gantt_chart(
                combined_df_filtered, unique_members, 
                data_inicio_analise, data_fim_analise, self.hoje
            )
            st.plotly_chart(fig, use_container_width=True)
            
            # Seção de conflitos
            st.write("---")
            st.header("Conflitos Detectados")
            
            conflitos_texto = DataProcessor.detect_conflicts(combined_df)
            
            if "Nenhum conflito detectado." in conflitos_texto:
                st.info(conflitos_texto)
            else:
                st.warning(f"```\n{conflitos_texto}\n```")
        else:
            st.warning("Não foi possível carregar os dados. Verifique se os arquivos Excel estão presentes e corretos.")


if __name__ == "__main__":
    app = App()
    app.run()