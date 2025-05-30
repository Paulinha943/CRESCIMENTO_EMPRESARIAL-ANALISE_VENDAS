# ================================================
# Script de Análise de Vendas 
# ================================================

import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import seaborn as sns

# ================================================
# Configuração Global dos Gráficos
# ================================================
sns.set(style="whitegrid", palette="pastel")
plt.rcParams['font.family'] = 'DejaVu Sans'  # Fonte compatível com Unicode

# ================================================
# Funções Auxiliares
# ================================================
def ajustar_colunas(df):
    """
    Padroniza os nomes das colunas.
    """
    df.columns = ['ID_LOJA', 'CIDADE', 'MES_ANO', 'PRODUTO', 'QTDE_VENDAS', 'VALOR']
    return df

def carregar_dados(caminho, aba):
    """
    Lê a planilha Excel especificada e retorna o DataFrame.
    """
    return pd.read_excel(caminho, sheet_name=aba)

def exibir_estatisticas(df):
    """
    Exibe estatísticas descritivas principais.
    """
    faturamento_total = df['VALOR'].sum()
    media_vendas = df['VALOR'].mean()
    mediana_vendas = df['VALOR'].median()
    moda_vendas = df['VALOR'].mode()

    print("="*70)
    print(f"FATURAMENTO TOTAL DA REDE: R$ {faturamento_total:,.2f}")
    print("="*70)
    print(f"Média Geral de Vendas: R$ {media_vendas:,.2f}")
    print(f"Mediana Geral de Vendas: R$ {mediana_vendas:,.2f}")
    if not moda_vendas.empty:
        for idx, valor in enumerate(moda_vendas.values, start=1):
            print(f"Moda {idx}: R$ {valor:,.2f}")
    else:
        print("Moda: Nenhuma")
    print("="*70)

    return faturamento_total, media_vendas, mediana_vendas, moda_vendas

def criar_grafico_barra(x, y, titulo, xlabel, ylabel, palette, orient='v', anotacao=False):
    """
    Cria um gráfico de barras com Seaborn.
    """
    plt.figure(figsize=(10, 6))
    if orient == 'v':
        sns.barplot(x=x, y=y, palette=palette)
        if anotacao:
            for idx, valor in enumerate(y):
                plt.text(idx, valor, f"R$ {valor:,.2f}", ha='center', va='bottom', fontsize=9, fontweight='bold')
    else:
        sns.barplot(x=x, y=y, palette=palette, orient='h')
        if anotacao:
            for idx, valor in enumerate(x):
                plt.text(valor, idx, f"R$ {valor:,.2f}", va='center', ha='left', fontsize=9, fontweight='bold')
    plt.title(titulo, fontsize=14, fontweight='bold')
    plt.xlabel(xlabel)
    plt.ylabel(ylabel)
    plt.tight_layout()
    plt.show()

# ================================================
# Execução Principal
# ================================================
def main():
    # 1. Carregar Dados
    arquivo = "Dados_ficticios_Analise de Dados.xlsx"
    df = carregar_dados(arquivo, "Sheet1")
    df = ajustar_colunas(df)

    # 2. Estatísticas Gerais
    exibir_estatisticas(df)

    # 3. Visualizações
    # ------------------------------------------------
    # (1) Faturamento Total por Loja
    faturamento_loja = df.groupby('ID_LOJA')['VALOR'].sum().sort_values(ascending=False)
    criar_grafico_barra(
        x=faturamento_loja.index,
        y=faturamento_loja.values,
        titulo="Faturamento Total por Loja",
        xlabel="ID Loja",
        ylabel="Faturamento (R$)",
        palette="Blues_d",
        orient='v',
        anotacao=True
    )

    # ------------------------------------------------
    # (2) Média de Vendas por Produto
    media_vendas_produto = df.groupby('PRODUTO')['VALOR'].mean().sort_values(ascending=False)
    criar_grafico_barra(
        x=media_vendas_produto.values,
        y=media_vendas_produto.index,
        titulo="Média de Vendas por Produto",
        xlabel="Média de Faturamento (R$)",
        ylabel="Produto",
        palette="Greens_d",
        orient='h',
        anotacao=True
    )

    # ------------------------------------------------
    # (3) Média de Vendas por Cidade
    media_vendas_cidade = df.groupby('CIDADE')['VALOR'].mean().sort_values(ascending=False)
    criar_grafico_barra(
        x=media_vendas_cidade.values,
        y=media_vendas_cidade.index,
        titulo="Média de Vendas por Cidade",
        xlabel="Média de Faturamento (R$)",
        ylabel="Cidade",
        palette="Oranges_d",
        orient='h',
        anotacao=True
    )

    # ------------------------------------------------
    # (4) Sazonalidade: Vendas ao Longo do Tempo
    df['MES_ANO'] = pd.to_datetime(df['MES_ANO'], format='%m/%Y')
    vendas_mensal = df.groupby('MES_ANO')['VALOR'].sum().reset_index()

    plt.figure(figsize=(12, 6))
    sns.lineplot(x='MES_ANO', y='VALOR', data=vendas_mensal, marker='o', color='purple', linewidth=2)
    plt.title("Vendas Totais por Mês/Ano (Sazonalidade)", fontsize=14, fontweight='bold')
    plt.xlabel("Mês/Ano")
    plt.ylabel("Faturamento (R$)")
    plt.xticks(rotation=45)
    plt.tight_layout()
    plt.show()

    # ------------------------------------------------
    # (5) Gráfico de Estatísticas Gerais (Média, Mediana e Moda)
    _, media, mediana, moda = exibir_estatisticas(df)
    estatisticas_dict = {
        'Média': media,
        'Mediana': mediana
    }
    if not moda.empty:
        for idx, valor in enumerate(moda.values, start=1):
            estatisticas_dict[f'Moda {idx}'] = valor
    else:
        estatisticas_dict['Moda'] = np.nan

    estatisticas_df = pd.DataFrame(list(estatisticas_dict.items()), columns=['Estatística', 'Valor'])

    plt.figure(figsize=(8, 5))
    sns.barplot(
        x='Valor',
        y='Estatística',
        data=estatisticas_df,
        hue='Estatística',       # Adicionado para compatibilidade futura
        palette='Set2',
        dodge=False,             # Garante barras não sobrepostas
        legend=False             # Remove legenda desnecessária
    )
    for index, row in estatisticas_df.iterrows():
        plt.text(row['Valor'], index, f"R$ {row['Valor']:,.2f}", va='center', ha='left',
                 fontsize=10, fontweight='bold', color='black')
    plt.title("Média, Mediana e Moda Geral da Rede", fontsize=14, fontweight='bold')
    plt.xlabel("Valor (R$)")
    plt.ylabel("")
    plt.tight_layout()
    plt.show()

# ================================================
# Execução
# ================================================
if __name__ == "__main__":
    main()
