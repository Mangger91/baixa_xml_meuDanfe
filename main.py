from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import logging
from selenium.webdriver.common.desired_capabilities import DesiredCapabilities
import openpyxl
import os
import time

# === CONFIGURAÇÕES ===

CAMINHO_EXCEL = r'C:\Scripts_e_Automacoes\MeuDanfe_NFe\Chaves_de_Acesso.xlsx'
PASTA_DESTINO = r'C:\Scripts_e_Automacoes\MeuDanfe_NFe\XML'
URL_SITE = 'https://meudanfe.com.br'
CHROME_PROFILE_PATH = r'C:\Users\mangger\AppData\Local\Google\Chrome\User Data\Profile 1'
TEMPO_TIMEOUT = 120  # segundos


def configurar_chromedriver():
    caminho_driver = r'C:\Scripts_e_Automacoes\MeuDanfe_NFe\chromedriver.exe'
    service = Service(executable_path=caminho_driver, log_path=os.devnull) 

    options = Options()
    options.add_argument("--log-level=3")
    options.add_argument("--disable-blink-features=AutomationControlled")
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-gpu")
    options.add_argument("--start-maximized")
    options.add_experimental_option("excludeSwitches", ["enable-automation"])
    options.add_experimental_option("useAutomationExtension", False)
    options.add_argument(f"user-data-dir={CHROME_PROFILE_PATH}")

    prefs = {
        "safebrowsing.enabled": True,
        "download.default_directory": PASTA_DESTINO,
        "download.prompt_for_download": False,
        "download.directory_upgrade": True,
        "safebrowsing.disable_download_protection": True
    }
    options.add_experimental_option("prefs", prefs)

    driver = webdriver.Chrome(options=options)
    driver.execute_script("Object.defineProperty(navigator, 'webdriver', {get: () => undefined})")
    return driver


def carregar_planilha():
    wb = openpyxl.load_workbook(CAMINHO_EXCEL)
    sheet = wb.active
    return wb, sheet


def baixar_xml(driver, wait, chave, row):
    try:
        print(f"🔎 Buscando chave: {chave}")
        search_box = wait.until(EC.element_to_be_clickable((By.ID, 'searchTxt')))
        search_box.clear()
        search_box.send_keys(chave)

        search_button = wait.until(EC.element_to_be_clickable((By.ID, 'searchBtn')))
        search_button.click()

        # Aguardar sté 60s para a página de resultado carregar
        try:
            esperar_carregamento_terminar(driver)
            wait.until(EC.presence_of_all_elements_located((By.ID, 'downloadXmlBtn'))) 
        except:
            print("❌ Página de resultado não carregou. Reiniciando...")    
            row[2].value = "ERRO"
            forcar_reinicio_site(driver)
            return  

        # Tenta clicar no botão de download
        esperar_carregamento_terminar(driver)
        download_button = wait.until(EC.element_to_be_clickable((By.ID, 'downloadXmlBtn')))
        download_button.click()
        time.sleep(5)  # Tempo para download iniciar

        # Verifica se o XML foi baixado
        caminho_arquivo = os.path.join(PASTA_DESTINO, f"{chave}.xml")
        status = "SUCESSO" if os.path.exists(caminho_arquivo) else "FALTANDO"
        print(f"📥 Resultado: {status}")
        row[2].value = status

    except Exception as e:
        print(f"❌ Erro ao processar a chave {chave}: {e}")
        row[2].value = "ERRO"

    finally:
        # Garante que voltou para a tela inicial
        try:
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "searchTxt")))
        except:
            print("↩️ Página travada. Redirecionando manualmente.")
            driver.get(URL_SITE)
            time.sleep(5)


def esperar_carregamento_terminar(driver, timeout=30):
    try:
        WebDriverWait(driver, timeout).until_not(
            EC.presence_of_element_located((By.CLASS_NAME, "jloading"))
        )
    except:
        print("⚠️ Timeout esperando o carregamento da página terminar.")

def forcar_reinicio_site(driver):
    print("🔁 Recarregando o site manualmente... ")
    driver.get(URL_SITE)
    time.sleep(5)
    esperar_carregamento_terminar(driver)


def main():
    driver = configurar_chromedriver()
    wait = WebDriverWait(driver, TEMPO_TIMEOUT)
    driver.get(URL_SITE)
    time.sleep(3)
    wb, sheet = carregar_planilha()

    for row in sheet.iter_rows(min_row=2, max_row=sheet.max_row, min_col=1, max_col=3):
        chave = row[0].value
        status = row[2].value

        if chave and (status is None or status == ""):
            baixar_xml(driver, wait, chave, row)

    wb.save(CAMINHO_EXCEL)
    wb.close()
    driver.quit()
    print("✅ Processo concluído com sucesso.")


if __name__ == "__main__":
    main()
