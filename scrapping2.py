import os
import time
import pandas as pd
import requests
from urllib.parse import urljoin, unquote

# Importações do Selenium (todas necessárias)
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from webdriver_manager.chrome import ChromeDriverManager

# --- CONFIGURAÇÕES (sem alterações) ---
NOME_ARQUIVO_PLANILHA = 'planilha_de_links.xlsx'
NOME_COLUNA_LINKS = 'Links'
NOME_COLUNA_PASTA = 'NomeDaPasta'
PASTA_DOWNLOADS_RAIZ = 'pdfs_baixados_selenium'
# -------------------------

def configurar_driver_chrome():
    # ... (sem alterações)
    chrome_options = Options()
    chrome_options.add_argument("--no-sandbox")
    chrome_options.add_argument("--disable-dev-shm-usage")
    servico = Service(ChromeDriverManager().install())
    driver = webdriver.Chrome(service=servico, options=chrome_options)
    return driver

def baixar_com_requests(driver, pasta_destino):
    # ... (sem alterações)
    wait = WebDriverWait(driver, 10) # Tempo de espera pode ser menor aqui
    xpath_final = "//a[@class='buttonLink' and contains(text(), 'Baixar')]"
    
    # ATENÇÃO: Mudança importante para lidar com múltiplos links na mesma página
    # Em vez de pegar um só, pegamos TODOS os links que correspondem
    links_finais = wait.until(EC.presence_of_all_elements_located((By.XPATH, xpath_final)))
    
    if not links_finais:
        # Se a lista estiver vazia, lança Timeout para ser capturado pela rotina
        raise TimeoutException("Nenhum link de 'Baixar' encontrado.")

    print(f"     ...encontrados {len(links_finais)} link(s) de 'Baixar'.")

    for link_element in links_finais:
        try:
            href_value = link_element.get_attribute('href')
            
            if 'document.location=' in href_value:
                relative_path = href_value.split("'")[1]
                url_arquivo = urljoin(driver.current_url, relative_path)
                
                if not os.path.exists(pasta_destino):
                    os.makedirs(pasta_destino)

                selenium_cookies = driver.get_cookies()
                requests_session = requests.Session()
                for cookie in selenium_cookies:
                    requests_session.cookies.set(cookie['name'], cookie['value'])

                resposta_pdf = requests_session.get(url_arquivo, stream=True, allow_redirects=True)
                resposta_pdf.raise_for_status()

                nome_arquivo = "arquivo.pdf"
                if "content-disposition" in resposta_pdf.headers:
                    nome_arquivo = resposta_pdf.headers['content-disposition'].split("filename=")[-1].strip('"')
                else:
                    nome_arquivo = unquote(url_arquivo.split('/')[-1].split('?')[0])

                caminho_salvar = os.path.join(pasta_destino, nome_arquivo)
                
                if os.path.exists(caminho_salvar):
                    print(f"        -- Arquivo '{nome_arquivo}' já existe. Pulando.")
                    continue

                print(f"        -- Baixando '{nome_arquivo}'...")
                with open(caminho_salvar, 'wb') as f:
                    for chunk in resposta_pdf.iter_content(chunk_size=8192):
                        f.write(chunk)
                print(f"        -- Salvo com sucesso.")
            else:
                print("     !! Link final com formato 'href' não reconhecido. Pulando este link.")
        except Exception as e:
            print(f"     !! Erro ao baixar um arquivo específico: {e}. Continuando...")
            continue # Continua para o próximo link na lista


# --- FUNÇÃO PRINCIPAL DE TRABALHO COM TESTE DE TODAS AS ROTINAS ---
def processar_url_com_rotinas(driver, url, pasta_destino):
    """
    Testa todas as rotinas de forma não-exclusiva para uma URL.
    """
    print(f"\n[+] Processando URL: {url}")
    print(f"  -> Destino: {pasta_destino}")
    
    # --- ROTINA 1: Download Direto ---
    try:
        print("  -> [Rotina 1] Testando download direto...")
        driver.get(url)
        baixar_com_requests(driver, pasta_destino)
        print("     ...Rotina 1 concluída.")
    except TimeoutException:
        print("     ...nenhum arquivo encontrado na Rotina 1.")
    except Exception as e:
        print(f"     ...erro inesperado na Rotina 1: {e}")

    # --- ROTINA 2: Clicar em "Anexos" ---
    try:
        print("  -> [Rotina 2] Testando aba 'Anexos'...")
        driver.get(url) # Recarrega a página para garantir um estado limpo
        wait = WebDriverWait(driver, 10)
        
        botao_anexos = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[contains(., 'Anexos')] | //a[contains(., 'Anexos')]")))
        botao_anexos.click()
        print("     ...clicou em 'Anexos'. Procurando arquivos.")
        
        baixar_com_requests(driver, pasta_destino)
        print("     ...Rotina 2 concluída.")
    except TimeoutException:
        print("     ...caminho da Rotina 2 não encontrado ou sem arquivos.")
    except Exception as e:
        print(f"     ...erro inesperado na Rotina 2: {e}")

    # --- ROTINA 3: Sequência Completa ---
    try:
        print("  -> [Rotina 3] Testando sequência completa...")
        driver.get(url) # Recarrega a página novamente
        wait = WebDriverWait(driver, 10)

        # Passo 3.1: "Plano de Trabalho"
        botao_plano = wait.until(EC.element_to_be_clickable((By.XPATH, "//*[contains(text(), 'Plano de Trabalho')]")))
        botao_plano.click()

        # Passo 3.2: "Anexos"
        botao_anexos_seq = wait.until(EC.element_to_be_clickable((By.XPATH, "//*[contains(text(), 'anexos')]")))
        botao_anexos_seq.click()
        
        # Passo 3.3: "Listar Anexos Proposta"
        botao_listar = wait.until(EC.element_to_be_clickable((By.XPATH, "//*[contains(text(), 'Listar Anexos Proposta')]")))
        botao_listar.click()
        print("     ...sequência completa executada. Procurando arquivos.")
        
        baixar_com_requests(driver, pasta_destino)
        print("     ...Rotina 3 concluída.")
    except TimeoutException:
        print("     ...caminho da Rotina 3 não encontrado ou sem arquivos.")
    except Exception as e:
        print(f"     ...erro inesperado na Rotina 3: {e}")

    print(f"[+] Finalizado o processamento para a URL: {url}")


def main():
    """Função principal que orquestra o processo"""
    # ... (código do main sem alterações, apenas chama a nova função)
    print("--- Iniciando Script de Download com Teste de Todas as Rotinas ---")
    driver = configurar_driver_chrome()
    
    try:
        df = pd.read_excel(NOME_ARQUIVO_PLANILHA)
        if NOME_COLUNA_LINKS not in df.columns or NOME_COLUNA_PASTA not in df.columns:
            print(f"!! ERRO: Verifique se as colunas '{NOME_COLUNA_PASTA}' e '{NOME_COLUNA_LINKS}' existem!")
            return

        for index, row in df.iterrows():
            # ... (código de leitura da linha sem alterações)
            nome_pasta = row[NOME_COLUNA_PASTA]
            url = row[NOME_COLUNA_LINKS]
            if pd.isna(url) or pd.isna(nome_pasta):
                continue
            
            nome_pasta_limpo = "".join(c for c in str(nome_pasta) if c.isalnum() or c in (' ', '_', '-')).rstrip()
            pasta_destino_final = os.path.join(PASTA_DOWNLOADS_RAIZ, nome_pasta_limpo)
            
            processar_url_com_rotinas(driver, url, pasta_destino_final)
    
    except FileNotFoundError:
        print(f"!! ERRO: Arquivo '{NOME_ARQUIVO_PLANILHA}' não encontrado.")
    except Exception as e:
        print(f"!! Ocorreu um erro inesperado no processo principal: {e}")
    finally:
        print("\n--- Fechando o navegador ---")
        driver.quit()
        
    print("--- Script finalizado! ---")


if __name__ == "__main__":
    main()
