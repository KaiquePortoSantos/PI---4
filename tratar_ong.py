import pandas as pd
import os
import numpy as np
import matplotlib.pyplot as plt

from datetime import datetime

# -------------------------------
# Ler dados brutos (todas as abas)
# -------------------------------
def ler_dados(caminho_entrada):
    print("📥 Lendo todas as abas do Excel...")
    xls = pd.ExcelFile(caminho_entrada)
    dfs = {}
    
    for sheet in xls.sheet_names:
        df = pd.read_excel(xls, sheet_name=sheet)
        print(f"Aba '{sheet}' carregada: {df.shape[0]} linhas, {df.shape[1]} colunas")
        dfs[sheet] = df
    
    return dfs

# -------------------------------
# Tratar dados
# -------------------------------
def tratar_dados(df):
    print("🧹 Tratando dados...")

    # Remover duplicados
    df = df.drop_duplicates()

    # Remover espaços extras antes de verificar linhas vazias
    df = df.applymap(lambda x: x.strip() if isinstance(x, str) else x)

    # Remover linhas completamente vazias
    df = df.dropna(how='all')

       # Preencher valores ausentes
    df.fillna({
        'coluna_num': 0,
        'coluna_texto': 'desconhecido',
        'data': pd.Timestamp('2000-01-01')
    }, inplace=True)

    # Padronizar formatos
    if 'data' in df.columns:
        df['data'] = pd.to_datetime(df['data'], errors='coerce')
    if 'coluna_texto' in df.columns:
        df['coluna_texto'] = df['coluna_texto'].str.lower().str.strip()

    # Criar coluna derivada
    if 'coluna1' in df.columns and 'coluna2' in df.columns:
        df['soma_colunas'] = df['coluna1'] + df['coluna2']

    print(f"✅ Dados tratados: {df.shape[0]} linhas, {df.shape[1]} colunas")
    return df

# -------------------------------
# Salvar todas as abas em um único arquivo
# -------------------------------
def salvar_todas_abas(dfs, pasta_saida, nome_arquivo=None):
    os.makedirs(pasta_saida, exist_ok=True)
    if nome_arquivo is None:
        nome_arquivo = f"ong_dados_limpos_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
    caminho_saida = os.path.join(pasta_saida, nome_arquivo)
    
    with pd.ExcelWriter(caminho_saida, engine='openpyxl') as writer:
        for nome_aba, df in dfs.items():
            if df.shape[0] == 0:
                print(f"⚠️ Pulando aba '{nome_aba}' pois está vazia")
                continue
            df_tratado = tratar_dados(df)
            # Limitar o nome da aba a 31 caracteres (limite do Excel)
            nome_aba_curto = nome_aba[:31]
            df_tratado.to_excel(writer, sheet_name=nome_aba_curto, index=False)
    
    print(f"💾 Arquivo final salvo: {caminho_saida}")
    

# ---------- Função de EDA ----------
def eda(df, pasta_graficos="graficos"):
    
    """Gera gráficos básicos de EDA"""
    os.makedirs(pasta_graficos, exist_ok=True)

    # Histogramas para colunas numéricas
    for coluna in df.select_dtypes(include=np.number).columns:
        plt.figure()
        df[coluna].hist(bins=20)
        plt.title(f'Distribuição: {coluna}')
        plt.xlabel(coluna)
        plt.ylabel('Frequência')
        plt.savefig(os.path.join(pasta_graficos, f'hist_{coluna}.png'))
        plt.close()

    # Séries temporais para coluna de data
    if 'data' in df.columns:
        plt.figure()
        df.groupby('data')['soma_colunas'].sum().plot()
        plt.title('Série temporal - soma_colunas')
        plt.xlabel('Data')
        plt.ylabel('Soma')
        plt.savefig(os.path.join(pasta_graficos, 'serie_temporal_soma_colunas.png'))
        plt.close()

    # Comparação de categorias
    for coluna in df.select_dtypes(include='object').columns:
        plt.figure()
        df[coluna].value_counts().plot(kind='bar')
        plt.title(f'Comparação de categorias: {coluna}')
        plt.xlabel(coluna)
        plt.ylabel('Frequência')
        plt.savefig(os.path.join(pasta_graficos, f'bar_{coluna}.png'))
        plt.close()

    print("EDA concluída. Gráficos salvos na pasta:", pasta_graficos)



# -------------------------------
# Execução do pipeline
# -------------------------------
if __name__ == "__main__":
    caminho_entrada = "ONG_dados_sinteticos.xlsx"
    pasta_saida = "dados_limpos"

    # Ler todas as abas
    dados_por_aba = ler_dados(caminho_entrada)

    # Salvar todas as abas tratadas
    salvar_todas_abas(dados_por_aba, pasta_saida)

    # Gerar gráficos apenas das atividades
    eda(dados_por_aba)



    print("🏁 Pipeline concluído com sucesso!")
