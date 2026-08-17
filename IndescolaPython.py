import numpy as np
import pandas as pd
from girth import twopl_mml, ability_eap

# ==============================================================================
# 1. TRATAMENTO E RECODIFICAÇÃO DOS MICRODADOS (ETL)
# ==============================================================================
def preparar_dados_censo(df_raw: pd.DataFrame) -> pd.DataFrame:
    """
    Filtra escolas públicas ativas e recodifica as variáveis do Censo em dummies (0 ou 1).
    """
    df = df_raw.copy()
    
    # Filtro de Escopo: Apenas escolas públicas em atividade
    df = df[(df['TP_DEPENDENCIA'].isin([1, 2, 3])) & (df['TP_SITUACAO_FUNCIONAMENTO'] == 1)].copy()
    
    # 1. Saneamento e Estrutura Básica
    df['TEMAGUARECODE'] = np.where(
        (df['IN_AGUA_POTAVEL'] == 1) | (df['IN_AGUA_REDE_PUBLICA'] == 1) | (df['IN_AGUA_POCO_ARTESIANO'] == 1), 1, 0
    )
    df['IN_ENERGIA_REDE_PUBLICA'] = np.where(df['IN_ENERGIA_REDE_PUBLICA'] == 1, 1, 0)
    df['TEMESGOTORECODE'] = np.where(
        (df['IN_ESGOTO_REDE_PUBLICA'] == 1) | (df['IN_ESGOTO_FOSSA_SEPTICA'] == 1), 1, 0
    )
    df['IN_LIXO_SERVICO_COLETA'] = np.where(df['IN_LIXO_SERVICO_COLETA'] == 1, 1, 0)
    
    # 2. Espaços Pedagógicos e Sociais
    df['BIBLIOOUSALADELEITURA'] = np.where(
        (df['IN_BIBLIOTECA'] == 1) | (df['IN_BIBLIOTECA_SALA_LEITURA'] == 1), 1, 0
    )
    df['IN_PROF_BIBLIOTECARIO'] = np.where(df['IN_PROF_BIBLIOTECARIO'] == 1, 1, 0)
    df['IN_QUADRA_ESPORTES'] = np.where(
        (df['IN_QUADRA_ESPORTES'] == 1) | (df['IN_QUADRA_ESPORTES_COBERTA'] == 1) | (df['IN_QUADRA_ESPORTES_DESCOBERTA'] == 1), 1, 0
    )
    df['IN_REFEITORIO'] = np.where(df['IN_REFEITORIO'] == 1, 1, 0)
    df['IN_SALA_PROFESSOR'] = np.where(df['IN_SALA_PROFESSOR'] == 1, 1, 0)
    
    # 3. Inclusão e Acessibilidade
    if 'TP_AEE' in df.columns:
        cond_aee = df['TP_AEE'].fillna(0).isin([1, 2])
    else:
        cond_aee = False

    df['IN_SALA_ATENDIMENTO_ESPECIAL'] = np.where(
        (df['IN_SALA_ATENDIMENTO_ESPECIAL'] == 1) | cond_aee, 1, 0
    )
    
    df['TEMACESSIBILIDADERECODE'] = np.where(
        (df['IN_BANHEIRO_PNE'] == 1) | (df['IN_ACESSIBILIDADE_RAMPAS'] == 1) |
        (df['IN_ACESSIBILIDADE_ELEVADOR'] == 1) | (df['IN_ACESSIBILIDADE_CORRIMAO'] == 1) |
        (df['IN_ACESSIBILIDADE_PISOS_TATEIS'] == 1) | (df['IN_ACESSIBILIDADE_VAO_LIVRE'] == 1), 1, 0
    )
    
    # 4. Conforto Térmico
    salas_util = df['QT_SALAS_UTILIZADAS'].fillna(0)
    salas_clima = df['QT_SALAS_UTILIZA_CLIMATIZADAS'].fillna(0)
    prop_clima = np.where(salas_util > 0, salas_clima / salas_util, 0)
    df['i_climgt30lt70'] = np.where((prop_clima > 0.30) & (prop_clima <= 0.70), 1, 0)
    df['i_climgt70'] = np.where(prop_clima > 0.70, 1, 0)
    
    # 5. Audiovisual e Tecnologia
    df['TEMEQUIPAMENTOSRECODE'] = np.where(
        (df['IN_EQUIP_SOM'] == 1) | (df['IN_EQUIP_TV'] == 1) | (df['IN_EQUIP_MULTIMIDIA'] == 1), 1, 0
    )
    df['i_tvgt03'] = np.where(df['QT_EQUIP_TV'].fillna(0) > 3, 1, 0)
    df['IN_INTERNET_APRENDIZAGEM'] = np.where(df['IN_INTERNET_APRENDIZAGEM'] == 1, 1, 0)
    
    # Computadores para Alunos
    comp_aluno = (
        df['QT_DESKTOP_ALUNO'].fillna(0) + 
        df['QT_COMP_PORTATIL_ALUNO'].fillna(0) + 
        df['QT_TABLET_ALUNO'].fillna(0)
    )
    df['COMPUTADORESPARAALUNOSRECODE'] = np.where(comp_aluno > 0, 1, 0)
    
    # 6. Variáveis Estruturais
    df['item_parque'] = np.where(df['IN_PARQUE_INFANTIL'] == 1, 1, 0)
    df['item_sanitario_ei'] = np.where(df['IN_BANHEIRO_EI'] == 1, 1, 0)
    df['item_lab_ciencias'] = np.where(df['IN_LABORATORIO_CIENCIAS'] == 1, 1, 0)
    df['item_lab_informatica'] = np.where(df['IN_LABORATORIO_INFORMATICA'] == 1, 1, 0)
    
    return df


# ==============================================================================
# 2. CÁLCULO TRI POR ETAPA (Com Padronização 50/10)
# ==============================================================================
def calcular_tri_etapa(df: pd.DataFrame, itens: list, col_id: str = 'CO_ENTIDADE') -> pd.DataFrame:
    df_valido = df.dropna(subset=itens, how='all').copy()
    if len(df_valido) < 50:
        return pd.DataFrame(columns=[col_id, 'score_etapa'])
    
    X_df = df_valido[itens].apply(pd.to_numeric, errors='coerce').fillna(0).astype(int)
    
    # Remoção de colunas sem variância (segurança estatística)
    cols_validas = [c for c in X_df.columns if X_df[c].nunique() > 1]
    if len(cols_validas) < 2:
        return pd.DataFrame(columns=[col_id, 'score_etapa'])
        
    # Transposição para o Girth (Itens x Respondentes)
    X = X_df[cols_validas].to_numpy().T
    
    # Calibração 2PL
    estimativas = twopl_mml(X)
    dificuldades = estimativas['Difficulty']
    discriminacoes = estimativas['Discrimination']
    
    # Habilidade Latente (EAP)
    theta = ability_eap(X, dificuldades, discriminacoes)
    
    # Padronização na Régua Normativa (Média 50, Desvio Padrão 10)
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


# ==============================================================================
# 3. PIPELINE PRINCIPAL E AGREGAÇÃO HIERÁRQUICA
# ==============================================================================
def executar_pipeline_indescola(caminho_csv_censo: str, pasta_saida: str = '.'):
    print("1. Carregando microdados do Censo Escolar...")
    df_raw = pd.read_csv(caminho_csv_censo, sep=';', encoding='latin1', low_memory=False)
    
    print("2. Aplicando regras de tratamento (ETL)...")
    df = preparar_dados_censo(df_raw)
    
    # Definição dos Grupos de Itens
    itens_base_tech = [
        'TEMAGUARECODE', 'IN_ENERGIA_REDE_PUBLICA', 'TEMESGOTORECODE', 'IN_LIXO_SERVICO_COLETA',
        'BIBLIOOUSALADELEITURA', 'IN_PROF_BIBLIOTECARIO', 'IN_QUADRA_ESPORTES', 'IN_REFEITORIO',
        'IN_SALA_PROFESSOR', 'IN_SALA_ATENDIMENTO_ESPECIAL', 'TEMACESSIBILIDADERECODE',
        'i_climgt30lt70', 'i_climgt70', 'TEMEQUIPAMENTOSRECODE', 'i_tvgt03',
        'COMPUTADORESPARAALUNOSRECODE', 'IN_INTERNET_APRENDIZAGEM'
    ]
    itens_base_infantil = [
        'TEMAGUARECODE', 'IN_ENERGIA_REDE_PUBLICA', 'TEMESGOTORECODE', 'IN_LIXO_SERVICO_COLETA',
        'BIBLIOOUSALADELEITURA', 'IN_QUADRA_ESPORTES', 'IN_REFEITORIO', 'IN_SALA_PROFESSOR',
        'IN_SALA_ATENDIMENTO_ESPECIAL', 'TEMACESSIBILIDADERECODE', 'i_climgt30lt70',
        'i_climgt70', 'TEMEQUIPAMENTOSRECODE', 'i_tvgt03'
    ]

    etapas_config = {
        'creche': {'flag': 'IN_INF_CRE', 'mat': 'QT_MAT_INF_CRE', 'itens': itens_base_infantil + ['item_parque', 'item_sanitario_ei']},
        'pre':    {'flag': 'IN_INF_PRE', 'mat': 'QT_MAT_INF_PRE', 'itens': itens_base_infantil + ['item_parque', 'item_sanitario_ei']},
        'efai':   {'flag': 'IN_FUND_AI', 'mat': 'QT_MAT_FUND_AI', 'itens': itens_base_tech},
        'efaf':   {'flag': 'IN_FUND_AF', 'mat': 'QT_MAT_FUND_AF', 'itens': itens_base_tech + ['item_lab_ciencias', 'item_lab_informatica']},
        'em':     {'flag': 'IN_MED',     'mat': 'QT_MAT_MED',     'itens': itens_base_tech + ['item_lab_ciencias', 'item_lab_informatica']}
    }
    
    print("3. Processando réguas da TRI para as 5 Etapas de Ensino...")
    
    # 1. Mapeamento das colunas de matrícula que devem ser preservadas
    cols_matriculas = [conf['mat'] for conf in etapas_config.values()]
    colunas_iniciais = ['CO_ENTIDADE', 'NO_ENTIDADE', 'SG_UF', 'CO_MUNICIPIO', 'NO_MUNICIPIO', 'TP_DEPENDENCIA'] + cols_matriculas
    
    # Inicializa df_final contendo as matrículas originais alinhadas
    df_final = df[colunas_iniciais].copy()
    for col in cols_matriculas:
        df_final[col] = df_final[col].fillna(0)
    
    # 2. Execução da TRI por Etapa e Merge Seguro por 'CO_ENTIDADE'
    for etapa, conf in etapas_config.items():
        sub_df = df[df[conf['flag']] == 1].copy()
        res_etapa = calcular_tri_etapa(sub_df, conf['itens']).rename(columns={'score_etapa': f'ind_{etapa}'})
        
        # O merge mantém a integridade do DataFrame
        df_final = df_final.merge(res_etapa, on='CO_ENTIDADE', how='left')

    # 3. Cálculo dos Acumuladores de Média Ponderada e Média Simples
    numerador_ponderado = 0
    denominador_ponderado = 0
    soma_notas_validas = 0
    qtd_etapas_validas = 0

    for etapa, conf in etapas_config.items():
        col_ind = f'ind_{etapa}'
        col_mat = conf['mat']
        
        # Identifica se a escola possui nota TRI válida na etapa
        tem_nota_valida = df_final[col_ind].notna()
        
        # Acumula os valores para a Média Ponderada
        numerador_ponderado += df_final[col_ind].fillna(0) * df_final[col_mat]
        denominador_ponderado += tem_nota_valida.astype(int) * df_final[col_mat]
        
        # Acumula os valores para a Média Simples (Fallback)
        soma_notas_validas += df_final[col_ind].fillna(0)
        qtd_etapas_validas += tem_nota_valida.astype(int)

    # 4. Cálculo do Índice Final com Regra de Regressão (Fallback)
    media_ponderada = np.where(denominador_ponderado > 0, numerador_ponderado / denominador_ponderado, np.nan)
    media_simples = np.where(qtd_etapas_validas > 0, soma_notas_validas / qtd_etapas_validas, np.nan)
    
    # Se a escola tem matrículas registradas nas etapas válidas, usa a Ponderada.
    # Se a escola tem nota na TRI mas a matrícula veio 0 no Censo, usa a Média Simples das etapas.
    df_final['Indice_entidade'] = np.where(denominador_ponderado > 0, media_ponderada, media_simples)
    df_final['TOTAL_MATRICULAS_ESCOLA'] = np.where(denominador_ponderado > 0, denominador_ponderado, qtd_etapas_validas)
    
    # Categorização nos Níveis Oficiais
    condicoes = [
        df_final['Indice_entidade'] < 40,
        df_final['Indice_entidade'] < 50,
        df_final['Indice_entidade'] < 60,
        df_final['Indice_entidade'] >= 60
    ]
    rotulos = ['Elementar', 'Básica', 'Adequada', 'Avançada']
    df_final['Nivel_Infraestrutura'] = np.select(condicoes, rotulos, default='Sem Escala')
    
    print("4. Agregando Índices para Redes Municipais e Estaduais...")
    df_validos = df_final.dropna(subset=['Indice_entidade']).copy()
    
    # Agregação Municipal (Ponderada por Matrículas)
    df_validos['peso_escola_mun'] = df_validos['Indice_entidade'] * df_validos['TOTAL_MATRICULAS_ESCOLA']
    
    df_municipios = df_validos.groupby(['SG_UF', 'CO_MUNICIPIO', 'NO_MUNICIPIO']).agg(
        Soma_Pesos_Mun=('peso_escola_mun', 'sum'),
        Soma_Matriculas_Mun=('TOTAL_MATRICULAS_ESCOLA', 'sum')
    ).reset_index()
    
    df_municipios['Ind_rede_municipal'] = df_municipios['Soma_Pesos_Mun'] / df_municipios['Soma_Matriculas_Mun']
    df_municipios = df_municipios.drop(columns=['Soma_Pesos_Mun'])
    
    # Agregação Estadual (Média Simples)
    df_estados = df_municipios.groupby('SG_UF').agg(
        Ind_rede_estadual=('Ind_rede_municipal', 'mean')
    ).reset_index()
    
    print("5. Exportando relatórios finais...")
    pasta_saida = 'resultados'
    df_final.to_csv(f"{pasta_saida}/Indescola_Escolas.csv", sep=';', index=False, encoding='utf-8-sig')
    df_municipios.to_csv(f"{pasta_saida}/Indescola_Municipios.csv", sep=';', index=False, encoding='utf-8-sig')
    df_estados.to_csv(f"{pasta_saida}/Indescola_Estados.csv", sep=';', index=False, encoding='utf-8-sig')
    
    print("=== PIPELINE CONCLUÍDO COM SUCESSO! ===")

# Para executar:
executar_pipeline_indescola(r'recursos\microdados_ed_basica_2023.csv')
