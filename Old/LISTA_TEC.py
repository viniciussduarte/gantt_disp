import os
import pandas as pd
from datetime import datetime
import getpass
import unicodedata
import tkinter as tk
from tkinter import messagebox
import warnings

warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")

def normalizar_string(s):
    return ''.join(c for c in unicodedata.normalize('NFD', s) if unicodedata.category(c) != 'Mn').lower().strip()

def ler_aba(xls, sheet_name, arquivo):
    try:
        df = pd.read_excel(xls, sheet_name=sheet_name)
        df['Arquivo'] = arquivo
        return df
    except Exception as e:
        print(f"Erro ao ler a aba '{sheet_name}' no arquivo {arquivo}: {e}")
        return None

def processar_arquivos(plataforma, arquivos_para_processar):
    usuario = getpass.getuser()
    data_atual = datetime.now().strftime("%d-%m-%Y")

    base_path = rf'C:\Users\{usuario}\PETROBRAS\BUZIOS PPO IPO IED - Documentos\11.Turbomáquinas\03. ATIVO 3\{plataforma}\99. Outros\02. Lista Técnica'
    destino = os.path.join(base_path, 'Fotografia Inventário e Lista Técnica')

    abas_alvo = {
        'inventario': {'nome_aba': 'Inventário', 'dados': [], 'colunas': None, 'saida': f'{data_atual}_Inventario.xlsx'},
        'lista tecnica': {'nome_aba': 'Lista Técnica', 'dados': [], 'colunas': None, 'saida': f'{data_atual}_Lista_Tecnica.xlsx'}
    }

    mensagens_erro = []

    for arquivo in os.listdir(base_path):
        if not (arquivo.endswith('.xlsx') and arquivo[:4] in arquivos_para_processar):
            continue

        try:
            xls = pd.ExcelFile(os.path.join(base_path, arquivo))
            abas = {normalizar_string(a): a for a in xls.sheet_names}

            for chave, info in abas_alvo.items():
                if chave in abas:
                    df = ler_aba(xls, info['nome_aba'], arquivo)
                    if df is not None:
                        if info['colunas'] is None:
                            info['colunas'] = df.columns.tolist()
                        elif df.columns.tolist() != info['colunas']:
                            colunas_esperadas = info['colunas']
                            colunas_encontradas = df.columns.tolist()
                            mensagens_erro.append(
                                f"Arquivo: {arquivo} (Aba: '{info['nome_aba']}')\n"
                                f"→ Colunas esperadas:\n  {colunas_esperadas}\n"
                                f"→ Colunas encontradas:\n  {colunas_encontradas}\n"
                            )
                        info['dados'].append(df)
        except Exception as e:
            print(f"Erro ao processar o arquivo {arquivo}: {e}")

    for chave, info in abas_alvo.items():
        if info['dados']:
            df_final = pd.concat(info['dados'], ignore_index=True)
            caminho_saida = os.path.join(destino, info['saida'])
            df_final.to_excel(caminho_saida, index=False)
            print(f"[{plataforma}] Arquivo '{info['nome_aba']}' salvo em: {caminho_saida}")
        else:
            print(f"[{plataforma}] Nenhuma aba '{info['nome_aba']}' encontrada.")

    if mensagens_erro:
        root = tk.Tk()
        root.withdraw()
        messagebox.showinfo("Erro nos Arquivos:", "\n\n".join(mensagens_erro))

if __name__ == "__main__":
    arquivos_para_processar = ["1225", "1227", "1231", "1238", "1252", "1254", "5147", "5336", "5412"]
    for plataforma in ["1. P-80", "2. P-82"]:
        processar_arquivos(plataforma, arquivos_para_processar)