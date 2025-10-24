import pandas as pd
import csv

def filtrar_planilhas(input_file, output_file):
    # Lê o arquivo CSV de entrada
    df = pd.read_csv(input_file)

    # Define as colunas relevantes
    col_auditor = 'AUDITOR'  # Coluna 4 no script original
    col_municipio = 'MUNICIPIO'  # Coluna 6 no script original
    col_unidade_adm = 'UNIDADE_ADMINISTRATINA'  # Coluna 8 no script original
    col_sigla = 'SIGLA_UNIDADE_ADM'  # Coluna 7 no script original

    # Cria uma lista para armazenar as linhas de saída
    output = []
    link_check = set()  # Conjunto para verificar links repetidos

    enunciados = df.iloc[0];

    # Itera sobre as linhas do DataFrame
    for index, row in df.iterrows():
        auditor = row[col_auditor]
        municipio = row[col_municipio]
        escola = row[col_unidade_adm]
        inep = row[col_sigla]
        unidade_administrativa = f"{escola}({inep})"

        if pd.isna(auditor):
            continue

        for col_index,col in enumerate(df.columns):
            cell_value = row[col]

            if isinstance(cell_value, str) and ('.jpg' in cell_value or '.jpeg' in cell_value):
                codigo = df.columns[df.columns.get_loc(col)]
                questao = enunciados[col]
                
                #adiciona comentário, caso exista
                comentario = row[df.columns[col_index + 1]] if col_index + 1 < len(df.columns) else ''
                comentario = comentario if pd.notna(comentario) else ''

                links_na_celula = [link.strip() for link in cell_value.split(',')]

                for link in links_na_celula:
                    if (link.endswith('.jpg') or link.endswith('.jpeg')) and link not in link_check:
                        output.append([
                            auditor, municipio, unidade_administrativa,
                            codigo, questao, comentario, '', '', link, '', ''
                        ])
                        link_check.add(link)

    # Cria o arquivo CSV de saída
    headers = ["Auditor", "Município", "Unidade_Administrativa", "Questão-Código",
               "Questão", "Comentário", "Tópico", "Subtópico", "Links", "IMAGEM", "Check"]

    with open(output_file, 'w', newline='', encoding='utf-8') as csvfile:
        writer = csv.writer(csvfile)
        writer.writerow(headers)
        writer.writerows(output)

    print(f"Arquivo de saída '{output_file}' criado com sucesso.")

# Uso da função
input_file = 'Base de Dados.csv'  # Substitua pelo nome do seu arquivo de entrada
output_file = 'Dados_Filtrados.csv'  # Nome do arquivo de saída

filtrar_planilhas(input_file, output_file)
