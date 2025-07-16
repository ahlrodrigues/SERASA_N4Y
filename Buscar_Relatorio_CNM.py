from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.chrome.options import Options
from dotenv import load_dotenv
import os
import time
import subprocess
import zipfile
import urllib.request
import shutil
import platform

# ==== Função para instalar ChromeDriver compatível ====

def instalar_chromedriver_compatível(driver_path="./chromedriver"):
    print("🔍 Verificando versão do Google Chrome instalada...")
    try:
        resultado = subprocess.run(
            ["/opt/google/chrome/chrome", "--version"],
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            text=True,
            check=True
        )
        versao_completa = resultado.stdout.strip().split()[-1]  # ex: "138.0.7204.100"
        versao_principal = versao_completa.split(".")[0]        # ex: "138"
        print(f"✅ Chrome instalado: versão {versao_completa}")
    except Exception as e:
        print("❌ Não foi possível obter a versão do Google Chrome:", e)
        return False

    print(f"🌐 Buscando ChromeDriver compatível com versão {versao_principal}...")

    sistema = platform.system().lower()
    if sistema != "linux":
        print("❌ Este instalador automático só foi testado em Linux.")
        return False

    try:
        url_zip = f"https://storage.googleapis.com/chrome-for-testing-public/{versao_completa}/linux64/chromedriver-linux64.zip"
        caminho_zip = "chromedriver.zip"

        urllib.request.urlretrieve(url_zip, caminho_zip)
        print("📦 Download do ChromeDriver concluído.")

        with zipfile.ZipFile(caminho_zip, 'r') as zip_ref:
            zip_ref.extractall("chromedriver_temp")

        # Caminho corrigido da estrutura interna do zip
        novo_driver = os.path.join("chromedriver_temp", "chromedriver-linux64", "chromedriver")

        if not os.path.exists(novo_driver):
            raise FileNotFoundError(f"Arquivo não encontrado: {novo_driver}")

        shutil.move(novo_driver, driver_path)
        os.chmod(driver_path, 0o755)

        os.remove(caminho_zip)
        shutil.rmtree("chromedriver_temp")
        print("✅ ChromeDriver atualizado com sucesso.")
        return True

    except Exception as e:
        print("❌ Falha ao baixar ou substituir o ChromeDriver:", e)
        return False

# ==== Configurações iniciais ====

# Atualiza o ChromeDriver antes de iniciar o Selenium
driver_path = "./chromedriver"
instalar_chromedriver_compatível(driver_path)

# Carrega variáveis do .env
load_dotenv()
USUARIO = os.getenv("USUARIO_CNM")
SENHA = os.getenv("SENHA_CNM")

# Define pasta de download
diretorio_download = os.path.abspath("download")
os.makedirs(diretorio_download, exist_ok=True)

# Configurações do Chrome
chrome_options = Options()
chrome_options.add_experimental_option("prefs", {
    "download.default_directory": diretorio_download,
    "download.prompt_for_download": False,
    "download.directory_upgrade": True,
    "safebrowsing.enabled": True
})

# Inicializa o navegador
driver = webdriver.Chrome(service=Service(driver_path), options=chrome_options)
driver.maximize_window()
wait = WebDriverWait(driver, 20)

# ==== Função auxiliar para aguardar download ====
def aguardar_download(nome_parcial, pasta, timeout=60):
    print("⏳ Aguardando o download finalizar...")
    limite = time.time() + timeout
    arquivo_final = None

    while time.time() < limite:
        arquivos = os.listdir(pasta)
        for nome in arquivos:
            if nome_parcial in nome and nome.endswith(".crdownload"):
                print(f"⌛ Baixando... {nome}")
            elif nome_parcial in nome and not nome.endswith(".crdownload"):
                arquivo_final = os.path.join(pasta, nome)
                print(f"✅ Download completo: {arquivo_final}")
                return arquivo_final
        time.sleep(1)

    raise TimeoutError("❌ Tempo limite ao esperar o fim do download.")

# ==== Ações ====

driver.get("https://appv2.creditonamedida.com.br/logar")

campo_usuario = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, 'input[placeholder="Digite seu usuário"]')))
campo_usuario.clear()
campo_usuario.send_keys(USUARIO)

campo_senha = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, 'input[placeholder="Digite sua senha"]')))
campo_senha.clear()
campo_senha.send_keys(SENHA)

botao_entrar = wait.until(EC.element_to_be_clickable((By.XPATH, '//button[.//span[text()="Logar"]]')))
botao_entrar.click()

wait.until(EC.url_changes("https://appv2.creditonamedida.com.br/logar"))
print("✅ Login realizado com sucesso.")

print("➡️ Localizando o botão 'Relatórios'...")
menu_relatorios = wait.until(EC.element_to_be_clickable((By.XPATH, '//a[contains(text(), "Relatórios")]')))
menu_relatorios.click()
time.sleep(1)
menu_relatorios.click()
print("✅ Botão 'Relatórios' clicado duas vezes.")

print("⏳ Aguardando 5 segundos para o submenu 'Extratos' aparecer...")
time.sleep(5)

print("➡️ Tentando localizar e clicar no link 'Extratos'...")
try:
    link_extratos = wait.until(EC.element_to_be_clickable((By.XPATH, '//a[contains(@href, "/relatorio/extrato")]')))
    link_extratos.click()
    print("✅ Link 'Extratos' clicado com sucesso.")
except Exception as e:
    print("❌ Erro ao tentar clicar em 'Extratos':", e)

campo_data_inicial = wait.until(EC.presence_of_element_located((By.NAME, "dataInicial")))
campo_data_inicial.clear()
campo_data_inicial.send_keys("01012000")
print("📅 Data inicial preenchida.")

botao_pesquisar = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, 'button[title="Pesquisar"]')))
botao_pesquisar.click()
print("🔍 Botão 'Pesquisar' clicado.")
time.sleep(5)

# Força a remoção do arquivo antigo, se existir
relatorio_cnm = os.path.join(diretorio_download, "Relatorio_CNM.xlsx")
if os.path.exists(relatorio_cnm):
    try:
        os.remove(relatorio_cnm)
        print("🗑️ Arquivo antigo 'Relatorio_CNM.xlsx' removido antes do novo download.")
    except Exception as e:
        print(f"⚠️ Não foi possível remover o arquivo antigo: {e}")

print("⏳ Aguardando botão 'Excel' ficar clicável...")
try:
    botao_excel = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, 'button[title="Excel"]')))
    driver.execute_script("arguments[0].click();", botao_excel)
    print("📥 Botão 'Excel' clicado com sucesso.")
except Exception as e:
    print("❌ Não foi possível clicar no botão 'Excel':", e)
    with open("pagina_extrato.html", "w", encoding="utf-8") as f:
        f.write(driver.page_source)
    print("📄 HTML da página salvo como 'pagina_extrato.html' para análise.")

try:
    arquivo_original = aguardar_download(".xlsx", diretorio_download, timeout=60)
    novo_nome = os.path.join(diretorio_download, "Relatorio_CNM.xlsx")
    os.rename(arquivo_original, novo_nome)
    print("✅ Arquivo renomeado para:", novo_nome)
except Exception as e:
    print("❌ Erro ao renomear arquivo:", e)

driver.quit()
print("✅ Processo concluído.")
