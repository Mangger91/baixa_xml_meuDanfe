import chromedriver_autoinstaller
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import openpyxl
import time
import os

# Caminhos
caminho_excel = r'C:\Robo - Baixas de XML\Outputs\CHAVE DE ACESSO.xlsx'
pasta_destino = r'C:\Robo - Baixas de XML\Outputs'

# Instalar a versão específica do ChromeDriver compatível com a versão do Chrome
chrome_version = '127.0.6533.120'  # Versão do Chrome que você tem
chromedriver_autoinstaller.install(True)  # Força a instalação do ChromeDriver para a versão do Chrome

# Configurações do Chrome WebDriver
chrome_options = Options()
chrome_options.add_argument("--disable-blink-features=AutomationControlled")
chrome_options.add_argument("--no-sandbox")
chrome_options.add_argument("--disable-gpu")
chrome_options.add_argument("--start-maximized")
chrome_options.add_experimental_option("excludeSwitches", ["enable-automation"])
chrome_options.add_experimental_option('useAutomationExtension', False)

# Desabilitar avisos de download inseguro
prefs = {
    "safebrowsing.enabled": True,  # Desativa os avisos de navegação segura
    "download.default_directory": pasta_destino,  # Define o local de download padrão
    "download.prompt_for_download": False,  # Baixa sem perguntar onde salvar
    "download.directory_upgrade": True,
    "safebrowsing.disable_download_protection": True  # Desabilita proteção de download
}
chrome_options.add_experimental_option("prefs", prefs)

# Adicionar o caminho para o perfil do Chrome
chrome_profile_path = r'C:\Users\mangger\AppData\Local\Google\Chrome\User Data\Profile 1'  
chrome_options.add_argument(f"user-data-dir={chrome_profile_path}")

# Inicializar WebDriver
driver = webdriver.Chrome(options=chrome_options)

# Definir script para remover as variáveis do navegador que indicam automação
driver.execute_script("Object.defineProperty(navigator, 'webdriver', {get: () => undefined})")

# Carregar planilha
wb = openpyxl.load_workbook(caminho_excel)
sheet = wb.active

# Abrir o site
driver.get('https://meudanfe.com.br')

# Aguardar o site carregar
wait = WebDriverWait(driver, 10)

# Inicializar variáveis de controle
paused = False
running = True

# Etapa 1: Realizar o download de todas as chaves
for row in sheet.iter_rows(min_row=2, max_row=sheet.max_row, min_col=1, max_col=3):
    chave = row[0].value
    empresa = row[1].value
    status = row[2].value

    if chave is None or status == 'SUCESSO':
        continue

    # Escrever a chave no campo de busca
    search_box = wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="get-danfe"]/div/div/div[1]/div/div/div/input')))
    search_box.clear()
    search_box.send_keys(chave)

    # Clicar no botão de busca
    search_button = wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="get-danfe"]/div/div/div[1]/div/div/div/button')))
    search_button.click()
    time.sleep(5)

    # Verificar se a página de download foi carregada
    try:
        wait.until(EC.url_contains("https://meudanfe.com.br/ver-danfe"))
    except Exception as e:
        print(f"Erro ao tentar carregar a página de download para a chave {chave}: {e}")
        status_cell = sheet.cell(row=row[0].row, column=3)
        status_cell.value = 'ERRO'
        continue  # Pula para a próxima chave
    
    # Clicar na opção de download XML
    try:
        download_xml_button = wait.until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[1]/div/div[1]/div/div[2]/button[1]')))
        download_xml_button.click()
    except Exception as e:
        print(f"Erro ao tentar clicar no botão de download XML: {e}")
    time.sleep(5)
    
    # Clicar no botão de nova consulta
    nova_consulta = wait.until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[1]/div/div[1]/div/div[1]/button')))
    nova_consulta.click()
    time.sleep(5)

    # Verificar se o download foi concluído e atualizar a planilha
    arquivo_baixado = os.path.join(pasta_destino, f'{chave}.xml')

    status_cell = sheet.cell(row=row[0].row, column=3)
    if os.path.exists(arquivo_baixado):
        status_cell.value = 'SUCESSO'
    else:
        status_cell.value = 'FALTANDO'

    wb.save(caminho_excel)

# Encerrar WebDriver
driver.quit()

wb.close()

print("Verificação concluída, automação encerrada.")
