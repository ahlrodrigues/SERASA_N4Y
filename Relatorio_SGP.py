
import os
import time
import shutil
import zipfile
import socket
from dotenv import load_dotenv
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service

# ===================== CONFIGURAÇÕES INICIAIS =====================
load_dotenv()
LOGIN = os.getenv("USUARIO_SGP")
SENHA = os.getenv("SENHA_SGP")
driver_path = "./chromedriver"
nome_parcial = "cliente-"
nome_final = "Relatorio_SGP.xlsx"

download_dir = os.path.abspath("download")
if not os.path.exists(download_dir):
    os.makedirs(download_dir)
    print(f"[INFO] Pasta de download criada em: {download_dir}")
else:
    print(f"[INFO] Usando pasta de download: {download_dir}")

if not LOGIN or not SENHA:
    print("[ERRO] USUARIO_SGP ou SENHA_SGP não definidos no .env")
    exit(1)

def verificar_porta_localhost(porta, timeout=2):
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as sock:
        sock.settimeout(timeout)
        resultado = sock.connect_ex(('localhost', porta))
        return resultado == 0

if not verificar_porta_localhost(48589):
    print("[AVISO] Serviço local na porta 48589 não está ativo ou demorando para responder.")

def esperar_download_e_renomear():
    print("[INFO] Aguardando download do arquivo...")
    timeout = 600
    polling = 2
    tempo_inicio = time.time()

    while time.time() - tempo_inicio < timeout:
        arquivos = os.listdir(download_dir)
        em_progresso = [f for f in arquivos if f.endswith(".crdownload")]
        finalizados = [f for f in arquivos if f.startswith(nome_parcial) and f.endswith(".xlsx")]

        if em_progresso:
            print("[INFO] Download em andamento...")
        elif finalizados:
            original_path = os.path.join(download_dir, finalizados[0])
            destino_path = os.path.join(download_dir, nome_final)

            try:
                with zipfile.ZipFile(original_path, 'r') as zip_ref:
                    if zip_ref.testzip() is not None:
                        raise zipfile.BadZipFile("Erro ao ler conteúdo interno.")
            except zipfile.BadZipFile:
                print("[ERRO] Arquivo .xlsx baixado está corrompido. Aguardando novo download...")
                time.sleep(polling)
                continue

            if os.path.getsize(original_path) < 10240:
                print("[ERRO] Arquivo baixado é muito pequeno. Aguardando novo download...")
                time.sleep(polling)
                continue

            shutil.move(original_path, destino_path)
            print(f"[SUCESSO] Arquivo baixado e renomeado para: {destino_path}")
            return True

        time.sleep(polling)

    print("[ERRO] Tempo excedido esperando download.")
    return False

print("[INFO] Iniciando Chrome...")
service = Service(driver_path)
options = webdriver.ChromeOptions()
prefs = {"download.default_directory": download_dir}
options.add_experimental_option("prefs", prefs)
options.add_argument("--disable-extensions")
options.add_argument("--disable-gpu")
options.add_argument("--no-sandbox")
options.add_argument("--disable-software-rasterizer")
options.add_argument("--disable-dev-shm-usage")

driver = webdriver.Chrome(service=service, options=options)
driver.maximize_window()
wait = WebDriverWait(driver, 20)

try:
    print("[INFO] Acessando o SGP...")
    driver.get("https://sgp.net4you.com.br/admin/cliente/list/")

    print("[INFO] Realizando login...")
    campo_login = wait.until(EC.presence_of_element_located((By.ID, "id_username")))
    campo_senha = wait.until(EC.presence_of_element_located((By.NAME, "password")))
    botao_entrar = wait.until(EC.element_to_be_clickable((By.ID, "entrar")))

    campo_login.send_keys(LOGIN)
    campo_senha.send_keys(SENHA)
    botao_entrar.click()
    print("[INFO] Login concluído.")

    print("[INFO] Abrindo aba de TAGs...")
    aba_tag = wait.until(EC.element_to_be_clickable((By.ID, "ui-id-2")))
    aba_tag.click()

    print("[INFO] Selecionando TAG NEGATIVADO...")
    campo_tag = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, '.select2-search__field')))
    campo_tag.click()
    campo_tag.send_keys("NEGATIVADO")
    time.sleep(2)
    campo_tag.send_keys(Keys.ENTER)

    print("[INFO] Executando consulta...")
    botao_consulta = wait.until(EC.element_to_be_clickable((By.ID, "botao_consulta")))
    botao_consulta.click()

    print("[INFO] Aguardando resultados...")
    time.sleep(5)

    print("[INFO] Aguardando botão de exportação (até 1440s)...")
    try:
        botao_excel = WebDriverWait(driver, 1440).until(EC.element_to_be_clickable((By.ID, "idprintexcel")))
        botao_excel.click()
        print("[SUCESSO] Botão de exportação clicado.")
    except Exception as e:
        print(f"[ERRO] Timeout ao aguardar botão Excel: {e}")
        raise

    if not esperar_download_e_renomear():
        raise Exception("Download não finalizado corretamente.")

    input("[INFO] Pressione Enter para encerrar...")

except Exception as e:
    print(f"[ERRO] Ocorreu um erro: {e}")

finally:
    driver.quit()
    print("[INFO] Navegador encerrado.")
