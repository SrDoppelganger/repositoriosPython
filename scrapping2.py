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
NOME_COLUNA_LINKS = 'Links'
# A coluna de pastas não é mais necessária, mas não precisa ser removida da planilha.

# --- NOVA CONFIGURAÇÃO DE PASTA ÚNICA ---
# Define o caminho ABSOLUTO para a pasta de downloads.
# os.getcwd() pega a pasta atual onde o script está rodando.
PASTA_DOWNLOADS_UNICA = os.path.join(os.getcwd(), 'TODOS_OS_PDFS')
# ---------------------------------------------------

def configurar_driver_chrome(pasta_download):
    """Configura o Chrome para baixar arquivos automaticamente para uma pasta específica."""
    chrome_options = Options()
    
    # Preferências essenciais para o download automático via Selenium
    prefs = {
        "download.default_directory": pasta_download, # Define a pasta de download
        "download.prompt_for_download": False, # Desativa a pergunta "Onde salvar?"
        "download.directory_upgrade": True,
        "plugins.always_open_pdf_externally": True # MUITO IMPORTANTE: Evita que o Chrome abra o PDF no navegador
    }
    chrome_options.add_experimental_option("prefs", prefs)
    
    chrome_options.add_argument("--no-sandbox")
    chrome_options.add_argument("--disable-dev-shm-usage")
    
    servico = Service(ChromeDriverManager().install())
    driver = webdriver.Chrome(service=servico, options=chrome_options)
    return driver

def esperar_download_concluir(pasta_download, timeout_segundos=120):
    """
    Espera um download terminar monitorando a pasta por arquivos temporários '.crdownload'.
    """
    print("        -- Aguardando download terminar...")
    tempo_inicial = time.time()
    while True:
        # Procura por arquivos que o Chrome cria durante o download
        arquivos_incompletos = [f for f in os.listdir(pasta_download) if f.endswith('.crdownload')]
        if not arquivos_incompletos:
            # Uma pequena pausa para garantir que o arquivo foi totalmente escrito no disco
            time.sleep(2) 
            print("        -- Download concluído!")
            return True
        if time.time() - tempo_inicial > timeout_segundos:
            print("        !! Tempo de espera do download excedido. O arquivo pode estar corrompido ou o download falhou.")
            # Tenta limpar o arquivo .crdownload restante
            for f in arquivos_incompletos:
                try:
                    os.remove(os.path.join(pasta_download, f))
                except OSError:
                    pass
            return False
        time.sleep(1)

def clicar_e_baixar_via_selenium(driver, pasta_download):
    """
    Encontra todos os botões 'Baixar', clica em cada um e espera o download terminar.
    """
    wait = WebDriverWait(driver, 10)
    xpath_final = "//a[@class='buttonLink' and contains(text(), 'Baixar')]"
    
    try:
        links_finais = wait.until(EC.presence_of_all_elements_located((By.XPATH, xpath_final)))
    except TimeoutException:
        raise TimeoutException("Nenhum link de 'Baixar' encontrado nesta página/aba.")

    print(f"     ...encontrados {len(links_finais)} link(s) de 'Baixar'.")

    for i, link_element in enumerate(links_finais):
        print(f"     -> Clicando no link de download {i + 1}/{len(links_finais)}...")
        try:
            # Clica no botão para iniciar o download pelo navegador
            link_element.click()
            # Espera o download atual terminar antes de clicar no próximo
            esperar_download_concluir(pasta_download)
        except Exception as e:
            print(f"     !! Erro ao clicar ou esperar download do link {i + 1}: {e}")
            # Se um clique falhar, tenta ir para o próximo
            continue

def processar_url_com_rotinas(driver, url, pasta_download_unica):
    """
    Testa as 3 rotinas usando o método de clique do Selenium para download.
    """
    print(f"\n[+] Processando URL: {url}")
    
    # --- ROTINA 1: Download Direto ---
    try:
        print("  -> [Rotina 1] Testando download direto...")
        driver.get(url)
        clicar_e_baixar_via_selenium(driver, pasta_download_unica)
        print("     ...Rotina 1 concluída.")
    except TimeoutException:
        print("     ...nenhum arquivo encontrado na Rotina 1.")
    except Exception as e:
        print(f"     ...erro inesperado na Rotina 1: {e}")

    # --- ROTINA 2: Clique duplo ---
    try:
        print("  -> [Rotina 2] Testando clique duplo em 'Projeto Básico/Termo de Referência'...")
        driver.get(url)
        wait = WebDriverWait(driver, 10)
        
        xpath_span = "//span[contains(., 'Projeto Básico/Termo de Referência')]"
        primeiro_span = wait.until(EC.element_to_be_clickable((By.XPATH, xpath_span)))
        primeiro_span.click()
        time.sleep(1)
        segundo_span = wait.until(EC.element_to_be_clickable((By.XPATH, xpath_span)))
        segundo_span.click()
        
        clicar_e_baixar_via_selenium(driver, pasta_download_unica)
        print("     ...Rotina 2 concluída.")
    except TimeoutException:
        print("     ...caminho da Rotina 2 não encontrado ou sem arquivos.")
    except Exception as e:
        print(f"     ...erro inesperado na Rotina 2: {e}")

    # --- ROTINA 3: Sequência Completa ---
    try:
        print("  -> [Rotina 3] Testando sequência completa...")
        driver.get(url)
        wait = WebDriverWait(driver, 10)

        wait.until(EC.element_to_be_clickable((By.XPATH, "//*[contains(text(), 'Plano de Trabalho')]"))).click()
        wait.until(EC.element_to_be_clickable((By.XPATH, "//*[contains(text(), 'anexos')]"))).click()
        wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="form_submit"]'))).click()
        
        clicar_e_baixar_via_selenium(driver, pasta_download_unica)
        print("     ...Rotina 3 concluída.")
    except TimeoutException:
        print("     ...caminho da Rotina 3 não encontrado ou sem arquivos.")
    except Exception as e:
        print(f"     ...erro inesperado na Rotina 3: {e}")

    print(f"[+] Finalizado o processamento para a URL: {url}")

def main():
    print("--- Iniciando Script de Download 100% via Selenium ---")

    # Cria a pasta de downloads única se ela não existir
    if not os.path.exists(PASTA_DOWNLOADS_UNICA):
        os.makedirs(PASTA_DOWNLOADS_UNICA)
        print(f"Pasta de downloads criada em: {PASTA_DOWNLOADS_UNICA}")

    driver = configurar_driver_chrome(PASTA_DOWNLOADS_UNICA)
    
    try:
        df = pd.read_excel(NOME_ARQUIVO_PLANILHA)
        if NOME_COLUNA_LINKS not in df.columns:
            print(f"!! ERRO: A coluna '{NOME_COLUNA_LINKS}' não foi encontrada!")
            return

        for index, row in df.iterrows():
            url = row[NOME_COLUNA_LINKS]
            if pd.isna(url):
                continue
            
            processar_url_com_rotinas(driver, url, PASTA_DOWNLOADS_UNICA)

            delay = 3
            print(f"\n--- Pausando por {delay} segundos para evitar bloqueio ---\n")
            time.sleep(delay)
    
    except FileNotFoundError:
        print(f"!! ERRO: Arquivo '{NOME_ARQUIVO_PLANILHA}' não encontrado.")
    except Exception as e:
        print(f"!! Ocorreu um erro inesperado no processo principal: {e}")
    finally:
        print("\n--- Fechando o navegador ---")
        driver.quit()
        
    print(f"--- Script finalizado! Todos os arquivos foram salvos em '{PASTA_DOWNLOADS_UNICA}' ---")

if __name__ == "__main__":
    main()
