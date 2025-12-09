import os
import time
import pandas as pd
import requests
from urllib.parse import urljoin, unquote

# Importações do Selenium
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException # Importante para capturar erros de espera
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

# --- NOVA FUNÇÃO DE LÓGICA DUPLA ---
def baixar_pdf_logica_dupla(driver, url, pasta_destino):
    """
    Tenta baixar o PDF usando a nova sequência de cliques.
    Se falhar, tenta o método antigo (download direto).
    """
    print(f"\n[+] Processando URL: {url}")
    print(f"  -> Destino: {pasta_destino}")

    try:
        driver.get(url)
        wait = WebDriverWait(driver, 15) # Tempo de espera um pouco menor para a tentativa
        
        # --- TENTATIVA 1: Novo Método (Sequência de Cliques) ---
        print("  -> Tentando o novo método (sequência de cliques)...")
        
        # Passo 1: Clicar em "Plano de Trabalho"
        botao_plano = wait.until(EC.element_to_be_clickable((By.XPATH, "//*[contains(text(), 'Plano de Trabalho')]")))
        botao_plano.click()
        print("     ...Passo 1 OK: 'Plano de Trabalho'")

        # Passo 2: Clicar em "anexos"
        botao_anexos = wait.until(EC.element_to_be_clickable((By.XPATH, "//*[contains(text(), 'anexos')]")))
        botao_anexos.click()
        print("     ...Passo 2 OK: 'anexos'")
        
        # Passo 3: Clicar em "Listar Anexos Proposta"
        botao_listar = wait.until(EC.element_to_be_clickable((By.XPATH, "//*[contains(text(), 'Listar Anexos Proposta')]")))
        botao_listar.click()
        print("     ...Passo 3 OK: 'Listar Anexos Proposta'")
        
        # Se todos os passos funcionaram, o código continua e baixa o arquivo aqui
        print("  -> Sequência concluída. Procurando o link final de download...")
        baixar_com_requests(driver, pasta_destino) # Chama a função auxiliar de download
        
        print("  -> Sucesso com o novo método!")

    except TimeoutException:
        # Se a sequência de cliques falhou (TimeoutException é o erro mais comum aqui),
        # o script pula para cá.
        print("  -> Novo método falhou (como esperado para páginas antigas).")
        print("  -> Tentando o método antigo (download direto)...")
        
        try:
            # --- TENTATIVA 2: Método Antigo (Download Direto) ---
            # O driver já está na página correta (driver.get(url) já foi chamado)
            baixar_com_requests(driver, pasta_destino) # Tenta baixar diretamente
            print("  -> Sucesso com o método antigo!")

        except TimeoutException:
            # Se o método antigo também falhar
            print(f"!! ERRO FINAL: Nenhum método de download funcionou para a URL: {url}")
            print(f"   Verifique se o botão 'Baixar' ou a sequência de cliques está correta para esta página.")
        except Exception as e:
            print(f"!! Ocorreu um erro inesperado no método antigo: {e}")

    except Exception as e:
        print(f"!! Ocorreu um erro inesperado no novo método: {e}")


def baixar_com_requests(driver, pasta_destino):
    """
    Função auxiliar que contém a lógica de encontrar o link 'Baixar'
    e fazer o download com a biblioteca requests.
    """
    wait = WebDriverWait(driver, 20)
    xpath_final = "//a[@class='buttonLink' and contains(text(), 'Baixar')]"
    
    link_final_element = wait.until(EC.presence_of_element_located((By.XPATH, xpath_final)))
    href_value = link_final_element.get_attribute('href')
    
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
            print(f"     -- Arquivo '{nome_arquivo}' já existe. Pulando.")
            return

        print(f"     -- Baixando '{nome_arquivo}'...")
        with open(caminho_salvar, 'wb') as f:
            for chunk in resposta_pdf.iter_content(chunk_size=8192):
                f.write(chunk)
        
        print(f"     -- Salvo com sucesso em: {caminho_salvar}")
    else:
        print("     !! Link final de download com formato 'href' não reconhecido.")


def main():
    """Função principal (sem alterações)"""
    print("--- Iniciando Script de Download com Lógica Dupla ---")
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
            
            # Chama a nova função de lógica dupla
            baixar_pdf_logica_dupla(driver, url, pasta_destino_final)
    
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
