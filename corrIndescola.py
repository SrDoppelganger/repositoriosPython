import pandas as pd
from scipy.stats import pearsonr, spearmanr

# 1. Carregar os dois arquivos (ou utilizar os DataFrames já carregados em memória)
# Substitua pelos caminhos reais das suas planilhas
df_python = pd.read_csv(r'resultados\Indescola_Escolas.csv', sep=';')
df_colega = pd.read_excel(r'recursos\indescola23-planilha.xlsx', sheet_name="indescola23 - municipais", decimal=',')  # ou pd.read_csv(...)

# 1. Cruzamento das bases
df_comparacao = pd.merge(
    df_python[['CO_ENTIDADE', 'Indice_entidade']],
    df_colega[['CO_ENTIDADE', 'Ind_entidade']],
    on='CO_ENTIDADE',
    how='inner'
)

# 2. Tratamento de texto: substitui vírgula por ponto (caso sejam strings)
df_comparacao['Indice_Infraestrutura_Final'] = df_comparacao['Indice_entidade'].astype(str).str.replace(',', '.')
df_comparacao['Ind_entidade'] = df_comparacao['Ind_entidade'].astype(str).str.replace(',', '.')

# 3. CONVERSÃO FORÇADA PARA NÚMERO (Resolve o AttributeError)
# 'errors=coerce' converte qualquer texto inválido para NaN
df_comparacao['Indice_Infraestrutura_Final'] = pd.to_numeric(df_comparacao['Indice_Infraestrutura_Final'], errors='coerce')
df_comparacao['Ind_entidade'] = pd.to_numeric(df_comparacao['Ind_entidade'], errors='coerce')

# 4. Remove linhas com valores nulos resultantes da conversão
df_validos = df_comparacao.dropna(subset=['Indice_Infraestrutura_Final', 'Ind_entidade']).copy()

# 5. Cálculo das correlações
r_pearson, p_val_pearson = pearsonr(df_validos['Indice_Infraestrutura_Final'], df_validos['Ind_entidade'])
rho_spearman, p_val_spearman = spearmanr(df_validos['Indice_Infraestrutura_Final'], df_validos['Ind_entidade'])

print("=== RESULTADOS ===")
print(f"Total de registros numéricos válidos: {len(df_validos)}")
print(f"Pearson (r):   {r_pearson:.4f} (p-valor: {p_val_pearson:.4e})")
print(f"Spearman (rho): {rho_spearman:.4f} (p-valor: {p_val_spearman:.4e})")
