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

# ==== Configurações iniciais ====

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
driver_path = "./chromedriver"
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

# Acessa o site de login
driver.get("https://appv2.creditonamedida.com.br/logar")

# Preenche login
campo_usuario = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, 'input[placeholder="Digite seu usuário"]')))
campo_usuario.clear()
campo_usuario.send_keys(USUARIO)

campo_senha = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, 'input[placeholder="Digite sua senha"]')))
campo_senha.clear()
campo_senha.send_keys(SENHA)

# Clica no botão "Logar"
botao_entrar = wait.until(EC.element_to_be_clickable((By.XPATH, '//button[.//span[text()="Logar"]]')))
botao_entrar.click()

# Aguarda redirecionamento
wait.until(EC.url_changes("https://appv2.creditonamedida.com.br/logar"))
print("✅ Login realizado com sucesso.")

# Clica no menu "Relatórios"
print("➡️ Localizando o botão 'Relatórios'...")
menu_relatorios = wait.until(EC.element_to_be_clickable((By.XPATH, '//a[contains(text(), "Relatórios")]')))
menu_relatorios.click()
time.sleep(1)
menu_relatorios.click()
print("✅ Botão 'Relatórios' clicado duas vezes.")

# Aguarda submenu "Extratos"
print("⏳ Aguardando 5 segundos para o submenu 'Extratos' aparecer...")
time.sleep(5)

# Clica em "Extratos"
print("➡️ Tentando localizar e clicar no link 'Extratos'...")
try:
    link_extratos = wait.until(EC.element_to_be_clickable((By.XPATH, '//a[contains(@href, "/relatorio/extrato")]')))
    link_extratos.click()
    print("✅ Link 'Extratos' clicado com sucesso.")
except Exception as e:
    print("❌ Erro ao tentar clicar em 'Extratos':", e)

# Preenche a data inicial
campo_data_inicial = wait.until(EC.presence_of_element_located((By.NAME, "dataInicial")))
campo_data_inicial.clear()
campo_data_inicial.send_keys("01012000")
print("📅 Data inicial preenchida.")

# Clica em "Pesquisar"
botao_pesquisar = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, 'button[title="Pesquisar"]')))
botao_pesquisar.click()
print("🔍 Botão 'Pesquisar' clicado.")
time.sleep(5)

# Clica em "Excel"
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

# Aguarda e renomeia o arquivo
try:
    arquivo_original = aguardar_download(".xlsx", diretorio_download, timeout=60)
    novo_nome = os.path.join(diretorio_download, "Relatorio_CNM.xlsx")
    os.rename(arquivo_original, novo_nome)
    print("✅ Arquivo renomeado para:", novo_nome)
except Exception as e:
    print("❌ Erro ao renomear arquivo:", e)

# Fecha o navegador
driver.quit()
print("✅ Processo concluído.")
