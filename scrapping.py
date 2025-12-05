import os
import time
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager

# --- CONFIGURAÇÕES ---
# Nome do arquivo da sua planilha Excel
NOME_ARQUIVO_PLANILHA = 'planilha_de_links.xlsx'
# Nome da coluna que contém os links
NOME_COLUNA_LINKS = 'Links'
# Texto EXATO (ou parcial) do botão de download. Seja o mais específico possível.
# Exemplos: 'Baixar', 'Download PDF', 'Fazer download'
TEXTO_BOTAO_BAIXAR = 'Baixar'
# Pasta onde os PDFs serão salvos
PASTA_DOWNLOADS = os.path.join(os.getcwd(), 'pdfs_baixados_selenium')
# ---------------------

def configurar_driver_chrome(pasta_download):
    """Configura o driver do Chrome com opções para download automático."""
    chrome_options = Options()
    # Preferências para o Chrome
    prefs = {
        "download.default_directory": pasta_download, # Define a pasta de download
        "download.prompt_for_download": False, # Desativa a pergunta sobre onde salvar
        "download.directory_upgrade": True,
        "plugins.always_open_pdf_externally": True # Evita que o PDF abra no navegador
    }
    chrome_options.add_experimental_option("prefs", prefs)
    
    # Instala e configura o ChromeDriver automaticamente
    servico = Service(ChromeDriverManager().install())
    
    # Cria a instância do driver
    driver = webdriver.Chrome(service=servico, options=chrome_options)
    return driver

def baixar_pdf_com_clique(driver, url, texto_botao):
    """Navega até uma URL, clica no botão de download e aguarda."""
    print(f"\n[+] Visitando a página: {url}")
    try:
        driver.get(url)

        # Espera ATÉ 15 segundos para que o botão seja clicável
        # O XPath procura QUALQUER elemento (*) que contenha o texto do botão.
        # Isso é flexível e funciona para <button>, <a>, <span>, etc.
        xpath_botao = f"//*[contains(text(), '{texto_botao}')]"
        print(f"  -> Procurando pelo botão com o texto: '{texto_botao}'")
        
        wait = WebDriverWait(driver, 15)
        botao_download = wait.until(EC.element_to_be_clickable((By.XPATH, xpath_botao)))
        
        print("  -> Botão encontrado! Clicando para iniciar o download...")
        botao_download.click()
        
        # --- Lógica de Espera do Download ---
        # Esta é uma abordagem simples: apenas espera um tempo fixo.
        # Pode não ser ideal para arquivos muito grandes ou conexões lentas.
        # Veja a seção "Melhoria Avançada" abaixo para uma solução mais robusta.
        print("  -- Aguardando 30 segundos para o download terminar...")
        time.sleep(30)
        print("  -- Download provavelmente concluído.")

    except Exception as e:
        print(f"!! Erro ao processar a página {url}. O botão pode não existir ou a página demorou para carregar.")
        # print(f"   Detalhe do erro: {e}") # Descomente para ver o erro técnico

def main():
    """Função principal que lê a planilha e inicia o processo com Selenium."""
    print("--- Iniciando Script de Download com Selenium ---")

    # Cria a pasta de downloads se não existir
    if not os.path.exists(PASTA_DOWNLOADS):
        print(f"Criando pasta de destino: '{PASTA_DOWNLOADS}'")
        os.makedirs(PASTA_DOWNLOADS)
    
    # Configura e inicia o navegador
    driver = configurar_driver_chrome(PASTA_DOWNLOADS)
    
    try:
        df = pd.read_excel(NOME_ARQUIVO_PLANILHA)
        if NOME_COLUNA_LINKS not in df.columns:
            print(f"!! ERRO: A coluna '{NOME_COLUNA_LINKS}' não foi encontrada!")
            return

        for url in df[NOME_COLUNA_LINKS]:
            if pd.notna(url):
                baixar_pdf_com_clique(driver, url, TEXTO_BOTAO_BAIXAR)
    
    except FileNotFoundError:
        print(f"!! ERRO: Arquivo '{NOME_ARQUIVO_PLANILHA}' não encontrado.")
    except Exception as e:
        print(f"!! Ocorreu um erro inesperado: {e}")
    finally:
        # É MUITO importante fechar o navegador no final
        print("\n--- Fechando o navegador ---")
        driver.quit()
        
    print("--- Script finalizado! ---")

if __name__ == "__main__":
    main()
