#lembrar de fazer um arquivo .txt com as placas e uma pasta com as planilhas (cuidado com os cabeçalhos)

import pandas as pd
import os

# --- CONFIGURAÇÃO ---
# Nome da pasta que contém as planilhas do DETRAN
pasta_planilhas = 'Planilhas'

# Nome do arquivo de texto com as placas que você quer encontrar
arquivo_placas = 'veiculos.txt'

# Nome da coluna que contém as placas nas suas planilhas
nome_coluna_placa = 'PLACA'

# Nome do arquivo de saída que será gerado com o resultado
arquivo_saida = 'resultado_auditoria.xlsx'
# --------------------


def auditar_placas_veiculos():
    """
    Função principal que lê as placas alvo, percorre as planilhas do DETRAN,
    encontra as correspondências e salva o resultado em uma nova planilha.
    """
    print("Iniciando o processo de auditoria...")

    # --- PASSO 1: Carregar a lista de placas para buscar ---
    try:
        with open(arquivo_placas, 'r', encoding='utf-8') as f:
            # Usamos .strip() para remover espaços em branco e quebras de linha
            # Usamos um set() para uma busca mais rápida e para evitar placas duplicadas
            placas_alvo = {linha.strip().upper() for linha in f if linha.strip()}
        print(f"Sucesso: {len(placas_alvo)} placas carregadas do arquivo '{arquivo_placas}'.")
    except FileNotFoundError:
        print(f"ERRO: O arquivo '{arquivo_placas}' não foi encontrado! Verifique o nome e o local do arquivo.")
        return # Encerra o script se o arquivo de placas não existir

    # --- PASSO 2: Percorrer as planilhas e encontrar as placas ---
    lista_dfs_encontrados = [] # Lista para guardar os pedaços de dados encontrados

    # Verifica se a pasta com as planilhas existe
    if not os.path.isdir(pasta_planilhas):
        print(f"ERRO: A pasta '{pasta_planilhas}' não foi encontrada! Verifique o nome e o local da pasta.")
        return

    # Lista todos os arquivos na pasta de planilhas
    for nome_arquivo in os.listdir(pasta_planilhas):
        # Processa apenas arquivos Excel (.xlsx ou .xls)
        if nome_arquivo.endswith(('.xlsx', '.xls')):
            caminho_completo = os.path.join(pasta_planilhas, nome_arquivo)
            print(f"Processando arquivo: {nome_arquivo}...")

            try:
                # Lê a planilha atual para um DataFrame do pandas
                df = pd.read_excel(caminho_completo)

                # Garante que a coluna de placas exista no arquivo
                df.columns = [col.upper() for col in df.columns]
                
                if nome_coluna_placa not in df.columns:
                    print(f"  Aviso: A coluna '{nome_coluna_placa}' não foi encontrada em '{nome_arquivo}'. Pulando este arquivo.")
                    continue

                # Garante que a coluna de placas seja do tipo string para comparação
                df[nome_coluna_placa] = df[nome_coluna_placa].astype(str).str.upper()

                # Filtra o DataFrame, mantendo apenas as linhas cujas placas estão na nossa lista alvo
                df_encontrados = df[df[nome_coluna_placa].isin(placas_alvo)]

                # Se encontrarmos alguma correspondência...
                if not df_encontrados.empty:
                    print(f"  > {len(df_encontrados)} placa(s) encontrada(s) neste arquivo.")
                    # Adiciona a nova coluna 'ARQUIVO' com o nome do arquivo de origem
                    df_encontrados = df_encontrados.copy() # Evita avisos de 'SettingWithCopyWarning'
                    df_encontrados['ARQUIVO'] = nome_arquivo
                    
                    # Adiciona este DataFrame com os resultados à nossa lista
                    lista_dfs_encontrados.append(df_encontrados)

            except Exception as e:
                print(f"  ERRO ao processar o arquivo '{nome_arquivo}': {e}")

    # --- PASSO 3: Juntar todos os resultados e salvar em um único arquivo ---
    if not lista_dfs_encontrados:
        print("\nAuditoria finalizada. Nenhuma das placas alvo foi encontrada nas planilhas.")
        return

    print("\nJuntando todos os resultados encontrados...")
    # Concatena todos os DataFrames da lista em um só
    df_final = pd.concat(lista_dfs_encontrados, ignore_index=True)

    # Salva o DataFrame final em um novo arquivo Excel
    try:
        df_final.to_excel(arquivo_saida, index=False)
        print(f"\nSucesso! Auditoria concluída. O resultado foi salvo em '{arquivo_saida}'.")
    except Exception as e:
        print(f"\nERRO ao salvar o arquivo de resultado: {e}")


# Executa a função principal quando o script é rodado
if __name__ == "__main__":

    auditar_placas_veiculos()

