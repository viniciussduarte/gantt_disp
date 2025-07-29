import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.dates as mdates
import warnings
import seaborn as sns
import streamlit as st
from datetime import datetime, timedelta

# --- Constants ---
FILE_PATH_ESTALEIRO = 'Planejamento Estaleiro.xlsx'
FILE_PATH_FERIAS = 'Férias.xlsx'
FILE_PATH_GERAL = 'Planejamento Geral.xlsx'
#FILE_PATH_ESTALEIRO = "https://petrobrasbr.sharepoint.com/:x:/t/bdoc_BUZIOS-PPO-IPO-IED/ETsMJONCHsJBrQQ62uUY5nkBitOBjDPJwu6oTy4Gkb51zQ?e=T1l3Hs"
#FILE_PATH_FERIAS = "https://petrobrasbr.sharepoint.com/:x:/t/bdoc_BUZIOS-PPO-IPO-IED/EbFHSF34ZGxJhfn-irMf43IB0AY0g3sRfva58HzjldztVg?e=yKzFlO"
#FILE_PATH_GERAL = "https://petrobrasbr.sharepoint.com/:x:/t/bdoc_BUZIOS-PPO-IPO-IED/EYibrfyUXJZLuwloRxKGunwBLJAM9KdJ4OZBb360T-xzSw?e=F9fBmj"

# --- Matplotlib and Seaborn Configuration (remains largely the same) ---
sns.set_style("whitegrid")
plt.rcParams['figure.figsize'] = (16, 8)
plt.rcParams['font.size'] = 10
plt.rcParams['text.usetex'] = False
plt.rcParams['mathtext.fontset'] = 'dejavusans'

# Suppress warnings
warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")

hoje = pd.Timestamp.today().normalize()

# --- Data Reading Functions (modified to remove Tkinter dependencies) ---

@st.cache_data
def read_estaleiro_data():
    try:
        # Leitura da aba "Equipe"
        equipe_df = pd.read_excel(FILE_PATH_ESTALEIRO, sheet_name='Equipe')
        equipe_df = equipe_df.iloc[:, [0, 1, 3, 4, 5, 7]].dropna(subset=[equipe_df.columns[4]])
        equipe_df.columns = ['Disciplina', 'Matrícula', 'Função', 'Projeto', 'Experiência', 'Nome']

        # Leitura da aba "Planejamento IED"
        planejamento_df = pd.read_excel(FILE_PATH_ESTALEIRO, sheet_name='Planejamento IED', skiprows=8, usecols="C:E")
        planejamento_df = planejamento_df.dropna(subset=[planejamento_df.columns[0]])
        planejamento_df.columns = ['Nome', 'Início', 'Término']

        # Converter datas para formato datetime
        planejamento_df['Início'] = pd.to_datetime(planejamento_df['Início'])
        planejamento_df['Término'] = pd.to_datetime(planejamento_df['Término'])

        # Merge com "Equipe" para trazer a matrícula
        planejamento_df = pd.merge(planejamento_df, equipe_df[['Nome', 'Matrícula', 'Disciplina', 'Função', 'Projeto']],
                                     on='Nome', how='left')
        planejamento_df['Tipo'] = 'Estaleiro'

        return equipe_df, planejamento_df
    except FileNotFoundError:
        st.error(f"Arquivo {FILE_PATH_ESTALEIRO} não encontrado. Por favor, verifique o caminho do arquivo.")
        return None, None

@st.cache_data
def read_ferias_data():
    try:
        df_ferias = pd.read_excel(FILE_PATH_FERIAS, skiprows=1, header=None)
        
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
        colunas_indices = [0, 1, 8, 9, 11, 12, 14, 15]
        df_selecionado = df_ferias.iloc[:, colunas_indices]
        df_selecionado.columns = [
            "Matrícula", "Nome do Empregado",
            "Primeira Parcela", "Termino Primeira Parcela",
            "Segunda Parcela", "Termino Segunda Parcela",
            "Terceira Parcela", "Termino Terceira Parcela"
        ]

        # Criar DataFrames para cada parcela
        df_primeira = df_selecionado[["Matrícula", "Nome do Empregado", "Primeira Parcela", "Termino Primeira Parcela"]].dropna(subset=["Primeira Parcela"])
        df_primeira = df_primeira.rename(columns={"Primeira Parcela": "Início", "Termino Primeira Parcela": "Término"})
        df_primeira["Tipo"] = "Férias"

        df_segunda = df_selecionado[["Matrícula", "Nome do Empregado", "Segunda Parcela", "Termino Segunda Parcela"]].dropna(subset=["Segunda Parcela"])
        df_segunda = df_segunda.rename(columns={"Segunda Parcela": "Início", "Termino Segunda Parcela": "Término"})
        df_segunda["Tipo"] = "Férias"

        df_terceira = df_selecionado[["Matrícula", "Nome do Empregado", "Terceira Parcela", "Termino Terceira Parcela"]].dropna(subset=["Terceira Parcela"])
        df_terceira = df_terceira.rename(columns={"Terceira Parcela": "Início", "Termino Terceira Parcela": "Término"})
        df_terceira["Tipo"] = "Férias"

        # Concatenar as parcelas de férias
        df_ferias_final = pd.concat([df_primeira, df_segunda, df_terceira], ignore_index=True)
        df_ferias_final['Início'] = pd.to_datetime(df_ferias_final['Início'])
        df_ferias_final['Término'] = pd.to_datetime(df_ferias_final['Término'])

        return df_ferias_final
    except FileNotFoundError:
        st.error(f"Arquivo {FILE_PATH_FERIAS} não encontrado. Por favor, verifique o caminho do arquivo.")
        return None

@st.cache_data
def read_planejamento_geral():
    try:
        planejamento_geral_df = pd.read_excel(FILE_PATH_GERAL, usecols=["Nome", "Matrícula", "Início", "Término", "DIAS", "Atividade"])
        planejamento_geral_df = planejamento_geral_df.rename(columns={'Atividade': 'Tipo'})
        planejamento_geral_df[['Início', 'Término']] = planejamento_geral_df[['Início', 'Término']].apply(pd.to_datetime)
        return planejamento_geral_df
    except FileNotFoundError:
        st.error(f"Arquivo {FILE_PATH_GERAL} não encontrado. Por favor, verifique o caminho do arquivo.")
        return None

# --- Helper Functions (remain largely the same) ---

def tem_conflito(row, combined_df, data_inicio_analise, data_fim_analise):
    periodos = combined_df[combined_df['Matrícula'] == row['Matrícula']]
    for _, periodo in periodos.iterrows():
        if (periodo['Início'] <= data_fim_analise and periodo['Término'] >= data_inicio_analise):
            return True
    return False

def detectar_conflitos(combined_df):
    """
    Detecta conflitos de agenda para todas as pessoas no DataFrame combined_df
    e retorna um texto formatado com os conflitos.
    
    Returns:
        str: Texto formatado com os conflitos detectados
    """
    # Lista para armazenar os conflitos
    conflitos_texto = []
    
    for matricula in combined_df['Matrícula'].unique():
        # Filtrar eventos desta pessoa
        eventos = combined_df[combined_df['Matrícula'] == matricula].copy()
        
        # Verificar se há pelo menos um evento
        if len(eventos) <= 1:
            continue
        
        # Obter o nome da pessoa (verificando se há eventos)
        if 'Nome' in eventos.columns and not eventos['Nome'].empty:
            nome = eventos['Nome'].iloc[0]  # Nome da pessoa
        else:
            nome = f"Matrícula: {matricula}"  # Usar matrícula se nome não disponível
        
        # Verificar conflitos (comparar cada par de eventos)
        conflito_encontrado = False
        for i in range(len(eventos)):
            for j in range(i+1, len(eventos)):
                evento1 = eventos.iloc[i]
                evento2 = eventos.iloc[j]
                
                # Verificar sobreposição de datas
                if (evento1['Início'] <= evento2['Término'] and evento1['Término'] >= evento2['Início']):
                    # Formatar texto do conflito
                    linha = f"{nome} - {evento1['Tipo']} ({evento1['Início'].strftime('%d/%m/%Y')} a {evento1['Término'].strftime('%d/%m/%Y')}) / {evento2['Tipo']} ({evento2['Início'].strftime('%d/%m/%Y')} a {evento2['Término'].strftime('%d/%m/%Y')})"
                    conflitos_texto.append(linha)
                    conflito_encontrado = True
                    break
            if conflito_encontrado:
                break
    
    # Retornar texto dos conflitos
    if conflitos_texto:
        return "CONFLITOS DETECTADOS:\n" + "\n".join(conflitos_texto)
    else:
        return "Nenhum conflito detectado."

def create_plot(combined_df, equipe_df, data_inicio_analise, data_fim_analise):
    fig, ax = plt.subplots(figsize=(16, 8))
    inicio = (hoje.replace(day=1) - pd.DateOffset(months=1))
    fim = inicio + pd.DateOffset(months=12) - timedelta(days=1)

    # Definir uma paleta de cores fixas para diferentes tipos de atividades
    cores_atividades = {
        'Estaleiro': 'blue',
        'Férias': 'red',
        'Workshop': 'green',
        'Treinamento': 'orange',
        'Embarque': 'purple',
        'Visita Técnica': 'teal',
        'Folga': 'magenta'
    }

    # Atribuir cores para tipos não especificados
    def atribuir_cor(tipo):
        if tipo in cores_atividades:
            return cores_atividades[tipo]
        else:
            # Se o tipo não estiver na lista, use uma cor padrão
            return 'gray'

    # Aplicar as cores aos tipos de atividades no DataFrame
    if not combined_df.empty and 'Tipo' in combined_df.columns:
        combined_df['Cor'] = combined_df['Tipo'].apply(atribuir_cor)

    # Ordenar todos os membros da equipe pela ordem desejada
    unique_members = equipe_df.sort_values(by=['Disciplina', 'Função', 'Projeto', 'Matrícula'], na_position='last').reset_index(drop=True)

    # Plotar os períodos de cada funcionário
    for i, row in unique_members.iterrows():
        matricula = row['Matrícula']
        periodos = combined_df[combined_df['Matrícula'] == matricula]
        for _, periodo in periodos.iterrows():
            cor = cores_atividades.get(periodo['Tipo'], 'brown')
            duracao = (periodo['Término'] - periodo['Início']).days
            if duracao >= 0:  # Verificação para evitar barras com largura negativa
                ax.barh(y=i, width=duracao + 1, left=periodo['Início'], height=0.4, color=cor)

        # Adicionar linha horizontal para separar cada nome
        ax.axhline(y=i, color='gray', linestyle=':', linewidth=0.5)

    # Adicionar linhas horizontais para separar disciplinas
    disciplinas = []
    disciplina_indices = {}
    
    # Identificar as disciplinas e seus índices finais
    disciplina_atual = None
    for i, row in unique_members.iterrows():
        disc = row['Disciplina']
        if pd.notna(disc) and (disciplina_atual is None or disc != disciplina_atual):
            disciplina_atual = disc
            disciplinas.append(disc)
            disciplina_indices[disc] = []
        
        if pd.notna(disc):
            disciplina_indices[disc].append(i)
    
    # Adicionar linhas e labels para cada disciplina
    for disc in disciplinas:
        if disc in disciplina_indices and disciplina_indices[disc]:
            # Encontrar o último índice para esta disciplina
            ultima_posicao = max(disciplina_indices[disc]) + 0.5
            ax.axhline(y=ultima_posicao, color='gray', linestyle='--', linewidth=1)
            
            # Adicionar rótulo da disciplina
            ax.text(
                pd.Timestamp(fim), 
                ultima_posicao - 0.25, 
                f"--- {disc} ---",
                ha='left',
                va='bottom',
                fontsize=10,
                fontweight='bold',
                color='gray'
            )

    # Configurar o eixo y com matrícula e nome
    y_labels = []
    y_colors = []
    for i, row in unique_members.iterrows():
        nome_completo = f"{row['Nome']} ({row['Projeto'] if pd.notna(row['Projeto']) else 'Sem Projeto'})"
        y_labels.append(nome_completo)
        # Pass combined_df to tem_conflito
        y_colors.append('green' if not tem_conflito(row, combined_df, data_inicio_analise, data_fim_analise) else 'black')

    ax.set_yticks(range(len(unique_members)))
    ax.set_yticklabels(y_labels)
    for label, color in zip(ax.get_yticklabels(), y_colors):
        label.set_color(color)

    ax.set_ylabel('Nome')
    ax.set_xlabel('Data')

    ax.axvline(x=hoje, color='red', linestyle='--', linewidth=2, label=' Hoje')
    ax.text(hoje, len(unique_members), ' Hoje', color='red', fontsize=12, verticalalignment='bottom')
    ax.axvspan(data_inicio_analise, data_fim_analise, color='gray', alpha=0.3, label='Período de Análise')
    ax.text(data_inicio_analise, len(unique_members), 'Análise', color='black', fontsize=12, verticalalignment='top')

    # Formatar o eixo x
    ax.xaxis.set_major_formatter(mdates.DateFormatter('%Y-%m'))
    ax.xaxis.set_major_locator(mdates.MonthLocator(interval=1))
    plt.xticks(rotation=45)
    ax.set_xlim(pd.Timestamp(inicio), pd.Timestamp(fim))

    # Criar legenda dinâmica com base nas cores utilizadas
    for tipo, cor in cores_atividades.items():
        ax.plot([], [], color=cor, label=tipo, linewidth=5)

    ax.legend(title='Legenda', bbox_to_anchor=(1.05, 1), loc='upper left')

    # Ajustar layout
    ax.set_title('Alocação da Equipe e Férias por Período (Ordenado por Disciplina, Função e Projeto)', fontweight='bold')
    for tick in ax.get_xticks():
        ax.axvline(x=tick, color='gray', linestyle='--', linewidth=0.5)
    
    # Detectar e adicionar conflitos como texto no final do gráfico
    conflitos_texto = detectar_conflitos(combined_df)
    fig.text(0.5, 0.01, conflitos_texto, ha='center', va='bottom', fontsize=9, 
             bbox=dict(boxstyle='round', facecolor='lightyellow', alpha=0.5))

    return fig

# --- Streamlit Application ---

def main_streamlit():
    st.set_page_config(layout="wide", page_title="Alocação de Equipe e Férias")
    st.title("Alocação de Equipe e Férias")

    st.sidebar.header("Configurações de Análise")
    
    # Streamlit date input for analysis period
    data_inicio_analise = st.sidebar.date_input("Data Início Análise", value=hoje - timedelta(days=30))
    data_fim_analise = st.sidebar.date_input("Data Fim Análise", value=hoje + timedelta(days=90))

    data_inicio_analise = pd.Timestamp(data_inicio_analise)
    data_fim_analise = pd.Timestamp(data_fim_analise)

    equipe_df, planejamento_df = read_estaleiro_data()
    df_ferias_final = read_ferias_data()
    planejamento_geral_df = read_planejamento_geral()

    if equipe_df is not None:
        dataframes_to_combine = []
        
        if planejamento_df is not None:
            dataframes_to_combine.append(planejamento_df[['Matrícula', 'Nome', 'Início', 'Término', 'Disciplina', 'Função', 'Projeto', 'Tipo']])
        
        if df_ferias_final is not None:
            dataframes_to_combine.append(df_ferias_final[['Matrícula', 'Nome do Empregado', 'Início', 'Término', 'Tipo']].rename(columns={'Nome do Empregado': 'Nome'}))
        
        if planejamento_geral_df is not None:
            dataframes_to_combine.append(planejamento_geral_df[['Matrícula', 'Nome', 'Início', 'Término', 'Tipo']])
        
        if dataframes_to_combine:
            combined_df = pd.concat(dataframes_to_combine, ignore_index=True)
        else:
            combined_df = pd.DataFrame(columns=['Matrícula', 'Nome', 'Início', 'Término', 'Disciplina', 'Função', 'Projeto', 'Tipo'])
        
        combined_df = combined_df.sort_values(by=['Disciplina', 'Função', 'Projeto', 'Matrícula'], na_position='last')
        
        matricula_mapping = {row['Matrícula']: row for _, row in equipe_df.iterrows()}
        
        for col in ['Disciplina', 'Função', 'Projeto']:
            if col not in combined_df.columns:
                combined_df[col] = None
        
        for idx, row in combined_df.iterrows():
            if pd.isna(row['Disciplina']) or pd.isna(row['Função']) or pd.isna(row['Projeto']):
                if row['Matrícula'] in matricula_mapping:
                    emp_data = matricula_mapping[row['Matrícula']]
                    combined_df.at[idx, 'Disciplina'] = emp_data['Disciplina']
                    combined_df.at[idx, 'Função'] = emp_data['Função']
                    combined_df.at[idx, 'Projeto'] = emp_data['Projeto']
        
        fig = create_plot(combined_df, equipe_df, data_inicio_analise, data_fim_analise)
        st.pyplot(fig) # Display the plot in Streamlit

    else:
        st.warning("Não foi possível carregar os dados. Verifique se os arquivos Excel estão presentes e corretos.")

if __name__ == "__main__":
    main_streamlit()