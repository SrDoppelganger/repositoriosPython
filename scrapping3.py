import os
import time
import pandas as pd

# Importações do Selenium
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from webdriver_manager.chrome import ChromeDriverManager

# --- CONFIGURAÇÕES ---
NOME_ARQUIVO_PLANILHA = 'planilha_de_links.xlsx'
NOME_COLUNA_LINKS = 'Links' # Coluna com as URLs
# --- NOVA CONFIGURAÇÃO DE ID ---
# Nome da sua "Coluna A" que contém o identificador (ex: '12345')
NOME_COLUNA_ID = 'ID_Arquivo'
# --------------------------------

PASTA_DOWNLOADS_UNICA = os.path.join(os.getcwd(), 'TODOS_OS_PDFS')

def configurar_driver_chrome(pasta_download):
    # ... (esta função permanece a mesma)
    chrome_options = Options()
    prefs = {
        "download.default_directory": pasta_download,
        "download.prompt_for_download": False,
        "download.directory_upgrade": True,
        "plugins.always_open_pdf_externally": True
    }
    chrome_options.add_experimental_option("prefs", prefs)
    chrome_options.add_argument("--no-sandbox")
    chrome_options.add_argument("--disable-dev-shm-usage")
    servico = Service(ChromeDriverManager().install())
    driver = webdriver.Chrome(service=servico, options=chrome_options)
    return driver

# --- NOVA FUNÇÃO PARA ESPERAR E RENOMEAR ---
def esperar_e_renomear_arquivo(pasta_download, id_arquivo, numero_arquivo, timeout_segundos=120):
    """
    Espera o download terminar e então renomeia o arquivo mais recente na pasta.
    """
    print("        -- Aguardando download e preparando para renomear...")
    tempo_inicial = time.time()
    
    # Loop de espera pelo fim do download (verificando arquivos .crdownload)
    while True:
        arquivos_incompletos = [f for f in os.listdir(pasta_download) if f.endswith('.crdownload')]
        if not arquivos_incompletos:
            time.sleep(2) # Pausa para garantir que o arquivo foi liberado pelo sistema
            break
        if time.time() - tempo_inicial > timeout_segundos:
            print("        !! Tempo de espera do download excedido.")
            return
        time.sleep(1)

    # Lógica para encontrar o arquivo mais recente e renomeá-lo
    try:
        # Pega todos os arquivos na pasta e encontra o mais novo pela data de modificação
        lista_de_arquivos = [os.path.join(pasta_download, f) for f in os.listdir(pasta_download)]
        if not lista_de_arquivos:
            print("        !! Pasta de downloads está vazia, não há o que renomear.")
            return

        caminho_antigo = max(lista_de_arquivos, key=os.path.getmtime)
        nome_antigo = os.path.basename(caminho_antigo)
        
        # Pega a extensão do arquivo original
        _, extensao = os.path.splitext(nome_antigo)
        
        # Constrói o novo nome do arquivo
        novo_nome = f"{id_arquivo}_{numero_arquivo}{extensao}"
        caminho_novo = os.path.join(pasta_download, novo_nome)
        
        # Renomeia o arquivo
        os.rename(caminho_antigo, caminho_novo)
        
        print(f"        -- Download concluído. Arquivo '{nome_antigo}' renomeado para '{novo_nome}'.")
        
    except Exception as e:
        print(f"        !! Erro ao tentar renomear o arquivo: {e}")

# --- FUNÇÃO DE CLIQUE MODIFICADA ---
def clicar_e_baixar_via_selenium(driver, pasta_download, id_arquivo):
    """
    Encontra todos os botões 'Baixar', clica em cada um e chama a função para renomear.
    """
    wait = WebDriverWait(driver, 10)
    xpath_final = "//a[@class='buttonLink' and contains(text(), 'Baixar')]"
    
    try:
        links_finais = wait.until(EC.presence_of_all_elements_located((By.XPATH, xpath_final)))
    except TimeoutException:
        raise TimeoutException("Nenhum link de 'Baixar' encontrado nesta página/aba.")

    print(f"     ...encontrados {len(links_finais)} link(s) de 'Baixar'.")

    for i, link_element in enumerate(links_finais):
        numero_do_arquivo = i + 1
        print(f"     -> Clicando no link de download {numero_do_arquivo}/{len(links_finais)}...")
        try:
            link_element.click()
            # Chama a nova função que espera e renomeia
            esperar_e_renomear_arquivo(pasta_download, id_arquivo, numero_do_arquivo)
        except Exception as e:
            print(f"     !! Erro durante o clique ou renomeação do arquivo {numero_do_arquivo}: {e}")
            continue

# --- FUNÇÃO DE ROTINAS MODIFICADA ---
def processar_url_com_rotinas(driver, url, pasta_download_unica, id_arquivo):
    """
    Testa as 3 rotinas, passando o id_arquivo para a função de download.
    """
    print(f"\n[+] Processando URL: {url} | ID: {id_arquivo}")
    
    # Rotina 1
    try:
        print("  -> [Rotina 1] Testando download direto...")
        driver.get(url)
        clicar_e_baixar_via_selenium(driver, pasta_download_unica, id_arquivo)
    except TimeoutException:
        print("     ...nenhum arquivo encontrado na Rotina 1.")
    except Exception as e:
        print(f"     ...erro inesperado na Rotina 1: {e}")

    # Rotina 2
    try:
        print("  -> [Rotina 2] Testando clique duplo...")
        driver.get(url)
        # ... (lógica da rotina 2 permanece a mesma) ...
        clicar_e_baixar_via_selenium(driver, pasta_download_unica, id_arquivo)
    except TimeoutException: # ...
    except Exception as e: # ...

    # Rotina 3
    try:
        print("  -> [Rotina 3] Testando sequência completa...")
        driver.get(url)
        # ... (lógica da rotina 3 permanece a mesma) ...
        clicar_e_baixar_via_selenium(driver, pasta_download_unica, id_arquivo)
    except TimeoutException: # ...
    except Exception as e: # ...

    print(f"[+] Finalizado o processamento para a URL: {url}")

def main():
    print("--- Iniciando Script com Renomeação Automática ---")

    if not os.path.exists(PASTA_DOWNLOADS_UNICA):
        os.makedirs(PASTA_DOWNLOADS_UNICA)

    driver = configurar_driver_chrome(PASTA_DOWNLOADS_UNICA)
    
    try:
        df = pd.read_excel(NOME_ARQUIVO_PLANILHA)
        # Verifica se as colunas essenciais existem
        if NOME_COLUNA_ID not in df.columns or NOME_COLUNA_LINKS not in df.columns:
            print(f"!! ERRO: Verifique se as colunas '{NOME_COLUNA_ID}' e '{NOME_COLUNA_LINKS}' existem na planilha!")
            return

        for index, row in df.iterrows():
            id_arquivo = row[NOME_COLUNA_ID]
            url = row[NOME_COLUNA_LINKS]
            
            if pd.isna(url) or pd.isna(id_arquivo):
                continue
            
            # Converte o ID para string para garantir que funcione no nome do arquivo
            id_arquivo_str = str(id_arquivo)
            
            processar_url_com_rotinas(driver, url, PASTA_DOWNLOADS_UNICA, id_arquivo_str)

            delay = 3
            print(f"\n--- Pausando por {delay} segundos ---\n")
            time.sleep(delay)
    
    except FileNotFoundError:
        print(f"!! ERRO: Arquivo '{NOME_ARQUIVO_PLANILHA}' não encontrado.")
    except Exception as e:
        print(f"!! Ocorreu um erro inesperado no processo principal: {e}")
    finally:
        print("\n--- Fechando o navegador ---")
        driver.quit()
        
    print(f"--- Script finalizado! ---")

if __name__ == "__main__":
    main()
