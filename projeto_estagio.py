import pandas as pd
from openpyxl import load_workbook  # Importe load_workbook a partir do openpyxl

# Carregue o arquivo de planilha existente
df = pd.read_excel("base_dados.xlsx")

# Calcule as taxas de validação e confirmação
df['Taxa de Validação'] = (df['Nº de vendas confirmadas'] + df['Nº de vendas canceladas']) / df['Nº total de vendas']
df['Taxa de Confirmação'] = df['Nº de vendas confirmadas'] / df['Nº total de vendas']

correlacao_validacao_reclamacoes = df['Taxa de Validação'].corr(df['Nº de reclamações por compra pendente'])
correlacao_confirmacao_reclamacoes = df['Taxa de Confirmação'].corr(df['Nº de reclamações por compra pendente'])
df["Correlação Validacao-Reclamacoes"] = correlacao_validacao_reclamacoes
df["Correlação Confirmacao-Reclamacoes"] = correlacao_confirmacao_reclamacoes
# Crie DataFrames para outras perguntas (certifique-se de que eles estejam corretamente calculados)
taxas_validacao_por_parceiro_mes_categoria = df.groupby(['Parceiro', 'Mês', 'Categoria'])['Taxa de Validação'].mean().reset_index()
taxas_confirmacao_por_parceiro_mes_categoria = df.groupby(['Parceiro', 'Mês', 'Categoria'])['Taxa de Confirmação'].mean().reset_index()
lojas_mais_reclamadas_por_mes_pendente = df.loc[df.groupby('Mês')['Nº de reclamações por compra pendente'].idxmax()].reset_index(drop=True)
meses_mais_reclamados = df.groupby('Mês')[['Nº de reclamações por compra pendente', 'Nº de reclamações por compra cancelada']].sum()

with pd.ExcelWriter("base_dados.xlsx", engine="openpyxl", mode="a") as writer:
    taxas_validacao_por_parceiro_mes_categoria.to_excel(writer, sheet_name="Taxas de Validação", index=False)
    taxas_confirmacao_por_parceiro_mes_categoria.to_excel(writer, sheet_name="Taxas de Confirmação", index=False)
    lojas_mais_reclamadas_por_mes_pendente.to_excel(writer, sheet_name="Lojas mais Reclamadas", index=False)
    meses_mais_reclamados.to_excel(writer, sheet_name="Meses mais Reclamados", index=False)


