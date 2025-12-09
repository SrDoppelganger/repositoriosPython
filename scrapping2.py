import os
import time
import pandas as pd
import requests
from urllib.parse import urljoin, unquote

# Importações do Selenium (essencial que estejam aqui)
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from webdriver_manager.chrome import ChromeDriverManager

# --- CONFIGURAÇÕES ---
NOME_ARQUIVO_PLANILHA = 'planilha_de_links.xlsx'
NOME_COLUNA_LINKS = 'Links'
NOME_COLUNA_PASTA = 'NomeDaPasta'
PASTA_DOWNLOADS_RAIZ = 'pdfs_baixados_selenium'
# -------------------------

def configurar_driver_chrome():
    """Configura e retorna uma instância do driver do Chrome."""
    chrome_options = Options()
    chrome_options.add_argument("--no-sandbox")
    chrome_options.add_argument("--disable-dev-shm-usage")
    # Para rodar sem abrir a janela do navegador, descomente a linha abaixo
    # chrome_options.add_argument("--headless")
    
    servico = Service(ChromeDriverManager().install())
    driver = webdriver.Chrome(service=servico, options=chrome_options)
    return driver

def executar_sequencia_e_baixar_pdf(driver, url, pasta_destino):
    """
    Executa uma sequência de cliques para chegar na página de download
    e então baixa o arquivo para a pasta de destino.
    """
    print(f"\n[+] Processando URL: {url}")
    print(f"  -> Destino: {pasta_destino}")

    try:
        driver.get(url)
        wait = WebDriverWait(driver, 20) # Espera de até 20 segundos

        # --- INÍCIO DA NOVA SEQUÊNCIA DE CLIQUES ---

        # Passo 1: Clicar em "Plano de Trabalho"
        print("  -> Passo 1: Clicando em 'Plano de Trabalho'...")
        # Usamos '//*' para ser genérico (pode ser <button>, <a>, <span>, etc.)
        botao_plano = wait.until(EC.element_to_be_clickable((By.XPATH, "//*[contains(text(), 'Plano de Trabalho')]")))
        botao_plano.click()

        # Passo 2: Clicar em "anexos"
        print("  -> Passo 2: Clicando em 'anexos'...")
        botao_anexos = wait.until(EC.element_to_be_clickable((By.XPATH, "//*[contains(text(), 'anexos')]")))
        botao_anexos.click()
        
        # Passo 3: Clicar em "Listar Anexos Proposta"
        print("  -> Passo 3: Clicando em 'Listar Anexos Proposta'...")
        botao_listar = wait.until(EC.element_to_be_clickable((By.XPATH, "//*[contains(text(), 'Listar Anexos Proposta')]")))
        botao_listar.click()

        # --- FIM DA NOVA SEQUÊNCIA ---
        
        # Agora, estamos na página final. Executamos a lógica de download que já tínhamos.
        print("  -> Na página final, procurando o botão 'Baixar'...")
        
        xpath_final = "//a[@class='buttonLink' and contains(text(), 'Baixar')]"
        link_final_element = wait.until(EC.presence_of_element_located((By.XPATH, xpath_final)))
        
        href_value = link_final_element.get_attribute('href')
        
        if 'document.location=' in href_value:
            relative_path = href_value.split("'")[1]
            url_arquivo = urljoin(driver.current_url, relative_path)
            
            print(f"  -> URL do arquivo encontrada: {url_arquivo}")

            # Cria a pasta de destino se não existir
            if not os.path.exists(pasta_destino):
                os.makedirs(pasta_destino)
                print(f"  -> Pasta '{pasta_destino}' criada.")

            # Lógica de download com 'requests'
            selenium_cookies = driver.get_cookies()
            requests_session = requests.Session()
            for cookie in selenium_cookies:
                requests_session.cookies.set(cookie['name'], cookie['value'])

            resposta_pdf = requests_session.get(url_arquivo, stream=True, allow_redirects=True)
            resposta_pdf.raise_for_status()

            nome_arquivo = "arquivo.pdf"
            if "content-disposition" in resposta_pdf.headers:
                header = resposta_pdf.headers['content-disposition']
                nome_arquivo = header.split("filename=")[-1].strip('"')
            else:
                nome_arquivo = unquote(url_arquivo.split('/')[-1].split('?')[0])

            caminho_salvar = os.path.join(pasta_destino, nome_arquivo)
            
            if os.path.exists(caminho_salvar):
                print(f"  -- Arquivo '{nome_arquivo}' já existe. Pulando.")
                return

            print(f"  -- Baixando '{nome_arquivo}'...")
            with open(caminho_salvar, 'wb') as f:
                for chunk in resposta_pdf.iter_content(chunk_size=8192):
                    f.write(chunk)
            
            print(f"  -- Salvo com sucesso em: {caminho_salvar}")
        else:
             print("  !! Link final de download com formato 'href' não reconhecido.")

    except Exception as e:
        print(f"!! Ocorreu um erro durante a sequência de cliques ou download para a URL {url}.")
        print(f"   Verifique se todos os botões existem na página ou se o tempo de espera é suficiente.")
        # print(f"   Detalhe técnico do erro: {e}") # Descomente para diagnóstico avançado

def main():
    """Função principal que lê a planilha e orquestra o processo."""
    print("--- Iniciando Script de Download com Sequência de Cliques ---")
    driver = configurar_driver_chrome()
    
    try:
        df = pd.read_excel(NOME_ARQUIVO_PLANILHA)
        if NOME_COLUNA_LINKS not in df.columns or NOME_COLUNA_PASTA not in df.columns:
            print(f"!! ERRO: Verifique se as colunas '{NOME_COLUNA_PASTA}' e '{NOME_COLUNA_LINKS}' existem!")
            return

        for index, row in df.iterrows():
            nome_pasta = row[NOME_COLUNA_PASTA]
            url = row[NOME_COLUNA_LINKS]
            if pd.isna(url) or pd.isna(nome_pasta):
                continue
            
            nome_pasta_limpo = "".join(c for c in str(nome_pasta) if c.isalnum() or c in (' ', '_', '-')).rstrip()
            pasta_destino_final = os.path.join(PASTA_DOWNLOADS_RAIZ, nome_pasta_limpo)
            
            # Chama a nova função que executa a sequência completa
            executar_sequencia_e_baixar_pdf(driver, url, pasta_destino_final)
    
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
