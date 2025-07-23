import undetected_chromedriver as uc
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
import openpyxl
import os
import time

# === CONFIGURAÇÕES ===

CAMINHO_EXCEL = r'C:\Scripts_e_Automacoes\MeuDanfe_NFe\Chaves_de_Acesso.xlsx'
PASTA_DESTINO = r'C:\Scripts_e_Automacoes\MeuDanfe_NFe\XML'
URL_SITE = 'https://meudanfe.com.br'
TEMPO_TIMEOUT = 120  # segundos

def configurar_chromedriver():
    options = uc.ChromeOptions()
    options.add_argument("--start-maximized")
    options.add_argument("--disable-blink-features=AutomationControlled")
    options.add_argument("--disable-gpu")
    options.add_argument("--disable-software-rasterizer")
    options.add_argument("--disable-accelerated-2d-canvas")
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")
    
    prefs = {
        "safebrowsing.enabled": True,
        "download.default_directory": PASTA_DESTINO,
        "download.prompt_for_download": False,
        "download.directory_upgrade": True
    }
    options.add_experimental_option("prefs", prefs)

    driver = uc.Chrome(options=options, use_subprocess=True)
    return driver

def carregar_planilha():
    wb = openpyxl.load_workbook(CAMINHO_EXCEL)
    sheet = wb.active
    return wb, sheet

def esperar_carregamento_terminar(driver, timeout=30):
    try:
        WebDriverWait(driver, timeout).until_not(
            EC.presence_of_element_located((By.CLASS_NAME, "jloading"))
        )
    except:
        print("⚠️ Timeout esperando o carregamento da página terminar.")

def baixar_xml(driver, wait, chave, row):
    try:
        print(f"🔎 Buscando chave: {chave}")
        search_box = wait.until(EC.element_to_be_clickable((By.ID, 'searchTxt')))
        search_box.clear()

        # Preenche a chave e envia com ENTER (alguns sites escutam esse evento)
        driver.execute_script(f"document.getElementById('searchTxt').value = '{chave}';")
        search_button = driver.find_element(By.ID, 'searchBtn')
        driver.execute_script("arguments[0].click();", search_button)

        # Espera inicial + aguarda botão de download aparecer
        time.sleep(2)
        WebDriverWait(driver, 30).until(
            EC.element_to_be_clickable((By.ID, 'downloadXmlBtn'))
        )

        # Clica no botão de download
        download_button = wait.until(EC.element_to_be_clickable((By.ID, 'downloadXmlBtn')))
        download_button.click()
        time.sleep(5)  # Tempo para download iniciar

        # Aguardar o download concluir
        caminho_final = os.path.join(PASTA_DESTINO, f"{chave}.xml")
        caminho_tempo = caminho_final + ".crdownload"
        tempo_maximo = 60  # Tempo máximo para esperar o download
        tempo_decorrido = 0
        
        while os.path.exists(caminho_tempo) and tempo_decorrido < tempo_maximo:
            time.sleep(1)
            tempo_decorrido += 1
        
        # Após aguardar, checa se o arquivo foi baixado corretamente
        if os.path.exists(caminho_final):
            status = "SUCESSO"
            print(f"✅ XML baixado com sucesso: {caminho_final}")
        else:
            status = "FALHA"
            print(f"❌ Falha ao baixar o XML para a chave: {chave}")
            
        print(f"Status do download: {status}")
        row[2].value = status

    except Exception as e:
        print(f"❌ Erro ao processar a chave {chave}: {e}")
        row[2].value = "ERRO"
        return driver, wait

    finally:
        # Voltar para tela inicial se travar
        try:
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "searchTxt")))
        except:
            print("↩️ Página travada. Redirecionando manualmente.")
            driver.get(URL_SITE)
            time.sleep(5)

    return driver, wait

def main():
    driver = configurar_chromedriver()
    driver.get(URL_SITE)
    wait = WebDriverWait(driver, TEMPO_TIMEOUT)
    time.sleep(3)
    wb, sheet = carregar_planilha()

    for row in sheet.iter_rows(min_row=2, max_row=sheet.max_row, min_col=1, max_col=3):
        chave = row[0].value
        status = row[2].value

        if chave and (status is None or status == ""):
            driver, wait = baixar_xml(driver, wait, chave, row)

    wb.save(CAMINHO_EXCEL)
    wb.close()
    driver.quit()
    print("✅ Processo concluído com sucesso.")

if __name__ == "__main__":
    main()
