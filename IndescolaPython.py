import numpy as np
import pandas as pd
from girth import twopl_mml,threepl_mml, ability_eap

# -------------------------------------------------------------------------
# 1. FUNÇÃO DE CÁLCULO TRI POR ETAPA (Régua 50/10)
# -------------------------------------------------------------------------
def calcular_tri_etapa(df: pd.DataFrame, itens: list, col_id: str = 'CO_ENTIDADE') -> pd.DataFrame:
    """
    Filtra os dados da etapa, ajusta modelo 2PL e retorna a nota padronizada (50, 10).
    """
    # Filtra apenas linhas que possuem ao menos uma resposta válida
    df_valido = df.dropna(subset=itens, how='all').copy()
    if len(df_valido) < 50:
        return pd.DataFrame(columns=[col_id, 'score_etapa'])
    
    X_df = df_valido[itens].apply(pd.to_numeric, errors='coerce').fillna(0).astype(int)
    
    # Filtro de variância
    cols_validas = [c for c in X_df.columns if X_df[c].nunique() > 1]
    if len(cols_validas) < 2:
        return pd.DataFrame(columns=[col_id, 'score_etapa'])
        
    # Transposição para o Girth: (Itens x Respondentes)
    X = X_df[cols_validas].to_numpy().T
    
    # Estimativa 2PL
    estimativas = threepl_mml(X)
    dificuldades = estimativas['Difficulty']
    discriminacoes = estimativas['Discrimination']
    
    # Estimação da Habilidade Latente (EAP)
    theta = ability_eap(X, dificuldades, discriminacoes)
    
    # Padronização na Régua 50/10
    std_theta = np.std(theta)
    if std_theta > 0:
        z_score = (theta - np.mean(theta)) / std_theta
        score_50_10 = 50 + (10 * z_score)
    else:
        score_50_10 = np.full_like(theta, 50.0)
        
    return pd.DataFrame({
        col_id: df_valido[col_id].values,
        'score_etapa': score_50_10
    })

# -------------------------------------------------------------------------
# 2. EXECUÇÃO DO PIPELINE
# -------------------------------------------------------------------------
def executar_pipeline_indescola(caminho_csv: str, caminho_saida: str):
    print("1. Lendo dados do Censo Escolar...")
    df_raw = pd.read_csv(caminho_csv, sep=';', encoding='latin1')
    
    # Exemplo simplificado de recodificação
    # (Adicione aqui todas as suas regras de binarização do Censo)
    df = df_raw.copy()
    df['TEMAGUARECODE'] = np.where((df['IN_AGUA_POTAVEL'] == 1) | (df['IN_AGUA_REDE_PUBLICA'] == 1), 1, 0)
    df['TEMESGOTORECODE'] = np.where((df['IN_ESGOTO_REDE_PUBLICA'] == 1) | (df['IN_ESGOTO_FOSSA_SEPTICA'] == 1), 1, 0)
    # ... recodificar demais variáveis ...

    # Definição dos grupos de itens por etapa
    itens_infantil = ['TEMAGUARECODE', 'TEMESGOTORECODE', 'IN_PARQUE_INFANTIL', 'IN_BANHEIRO_EI']
    itens_fundamental = ['TEMAGUARECODE', 'TEMESGOTORECODE', 'IN_BIBLIOTECA', 'IN_LABORATORIO_INFORMATICA']
    
    print("2. Calculando TRI por Etapa...")
    # Exemplo para Creche e Fundamental
    res_creche = calcular_tri_etapa(df[df['IN_INF_CRE'] == 1], itens_infantil).rename(columns={'score_etapa': 'ind_creche'})
    res_ef = calcular_tri_etapa(df[df['IN_FUND_AI'] == 1], itens_fundamental).rename(columns={'score_etapa': 'ind_ef'})
    
    print("3. Consolidando e Ponderando por Matrículas...")
    
    # 3.1 Definir todas as colunas de matrículas usadas no Censo (baseado no script R)
    colunas_matriculas = [
        'QT_MAT_INF_CRE',  # Creche
        'QT_MAT_INF_PRE',  # Pré-escola
        'QT_MAT_FUND_AI',  # Ensino Fundamental - Anos Iniciais
        'QT_MAT_FUND_AF',  # Ensino Fundamental - Anos Finais
        'QT_MAT_MED'       # Ensino Médio
    ]
    
    # 3.2 Selecionar as colunas base para o dataframe final
    colunas_base = ['CO_ENTIDADE', 'NO_ENTIDADE', 'SG_UF', 'CO_MUNICIPIO'] + colunas_matriculas
    df_final = df[colunas_base].copy()
    
    # Preencher valores nulos (NaN) com 0 nas colunas de matrícula para evitar erros de soma
    df_final[colunas_matriculas] = df_final[colunas_matriculas].fillna(0)
    
    # 3.3 Trazer os resultados calculados da TRI para o dataframe final
    # (Supondo que você calculou res_creche, res_pre, res_efai, res_efaf, res_em)
    df_final = df_final.merge(res_creche, on='CO_ENTIDADE', how='left')
    df_final = df_final.merge(res_ef, on='CO_ENTIDADE', how='left')
    # Adicione os merges das outras etapas aqui...
    
    # 3.4 Cálculo do Numerador e Denominador para Média Ponderada da Escola
    # O denominador soma apenas as matrículas das etapas que possuem índice (não são nulas)
    num = (df_final['ind_creche'].fillna(0) * df_final['QT_MAT_INF_CRE'] + 
           df_final['ind_ef'].fillna(0) * df_final['QT_MAT_FUND_AI']) # Adicione as outras etapas
    
    denominador_valido = (df_final['ind_creche'].notna().astype(int) * df_final['QT_MAT_INF_CRE'] + 
                          df_final['ind_ef'].notna().astype(int) * df_final['QT_MAT_FUND_AI']) # Adicione as outras etapas
    
    # Calcula o índice da escola
    df_final['Indice_Infraestrutura_Final'] = np.where(denominador_valido > 0, num / denominador_valido, np.nan)
    
    # -------------------------------------------------------------------------
    # NOVO: Criando a coluna de Total de Matrículas para a Agregação Municipal
    # -------------------------------------------------------------------------
    # Conforme o relatório técnico, para o índice municipal, usamos o total de matrículas da escola
    # como peso. Usaremos o 'denominador_valido' pois ele reflete exatamente os alunos 
    # que foram contemplados pelo cálculo do índice.
    
    df_final['TOTAL_MATRICULAS_ESCOLA'] = denominador_valido
    
    # Removemos escolas que ficaram sem índice (NaN) para não distorcer a agregação do município
    df_agregacao = df_final.dropna(subset=['Indice_Infraestrutura_Final']).copy()

    print("4. Iniciando Agregações (Municipal e Estadual)...")
    
    # Agregação Municipal (Ponderada pelas matrículas)
    df_agregacao['peso_escola_mun'] = df_agregacao['Indice_Infraestrutura_Final'] * df_agregacao['TOTAL_MATRICULAS_ESCOLA']
    
    df_municipios = df_agregacao.groupby(['SG_UF', 'CO_MUNICIPIO']).agg(
        Soma_Pesos_Mun=('peso_escola_mun', 'sum'),
        Soma_Matriculas_Mun=('TOTAL_MATRICULAS_ESCOLA', 'sum')
    ).reset_index()
    
    df_municipios['Ind_rede_municipal'] = df_municipios['Soma_Pesos_Mun'] / df_municipios['Soma_Matriculas_Mun']

    # Agregação Estadual (Média simples dos municípios)
    df_estados = df_municipios.groupby('SG_UF').agg(
        Ind_rede_estadual=('Ind_rede_municipal', 'mean')
    ).reset_index()
    
    print(f"4. Salvando resultado em {caminho_saida}...")
    df_final.to_csv(caminho_saida, index=False, sep=';', encoding='utf-8-sig')
    print("Concluído com sucesso!")

    # -------------------------------------------------------------------------
    # PASSO EXTRA: Agregação para Redes Municipais (Ind_rede_municipal)
    # Conforme exigido pela metodologia: ponderado pelo total de matrículas da escola
    # -------------------------------------------------------------------------
    # Supondo que df_final já possui 'Indice_Infraestrutura_Final' (Ind_entidade) 
    # e 'TOTAL_MATRICULAS_ESCOLA', 'CO_MUNICIPIO', 'SG_UF'

    # Calcula o peso absoluto de cada escola no município
    df_final['peso_escola_mun'] = df_final['Indice_Infraestrutura_Final'] * df_final['TOTAL_MATRICULAS_ESCOLA']

    # Agrupa por município
    df_municipios = df_final.groupby(['SG_UF', 'CO_MUNICIPIO']).agg(
        Soma_Pesos_Mun=('peso_escola_mun', 'sum'),
        Soma_Matriculas_Mun=('TOTAL_MATRICULAS_ESCOLA', 'sum')
    ).reset_index()

    # Calcula o índice municipal
    df_municipios['Ind_rede_municipal'] = df_municipios['Soma_Pesos_Mun'] / df_municipios['Soma_Matriculas_Mun']


    # -------------------------------------------------------------------------
    # PASSO EXTRA: Agregação para Redes Estaduais
    # Conforme exigido pela metodologia: média aritmética simples dos municípios
    # -------------------------------------------------------------------------
    df_estados = df_municipios.groupby('SG_UF').agg(
        Ind_rede_estadual=('Ind_rede_municipal', 'mean') # Média simples
    ).reset_index()

# Para rodar:
executar_pipeline_indescola('recursos/microdados_ed_basica_2023.csv', 'resultados/resultado_tri_python.csv')
