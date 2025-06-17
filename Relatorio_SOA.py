import os
import time
from dotenv import load_dotenv, dotenv_values
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

# ==== Configurações iniciais ====

dotenv_path = os.path.join(os.path.dirname(__file__), ".env")
load_dotenv(dotenv_path)
config = dotenv_values(dotenv_path)
print("[DEBUG] Conteúdo carregado do .env:", config)

LOGIN = config.get("USUARIO_SOA")
SENHA = config.get("SENHA_SOA")

if not LOGIN or not SENHA:
    print("[ERRO] Variáveis USUARIO_SOA ou SENHA_SOA não estão definidas no .env ou estão vazias.")
    print(f"[DEBUG] USUARIO_SOA = {LOGIN}")
    print(f"[DEBUG] SENHA_SOA = {'<vazio>' if not SENHA else '***'}")
    exit(1)

driver_path = "./chromedriver"
download_dir = os.path.abspath("download")
os.makedirs(download_dir, exist_ok=True)

chrome_options = Options()
chrome_options.add_experimental_option("prefs", {
    "download.default_directory": download_dir,
    "download.prompt_for_download": False,
    "directory_upgrade": True,
    "safebrowsing.enabled": True
})

driver = webdriver.Chrome(service=Service(driver_path), options=chrome_options)
driver.maximize_window()
wait = WebDriverWait(driver, 15)

def aguardar_download(pasta, timeout=120):
    print("[INFO] Aguardando novo arquivo .csv na pasta de download...")
    arquivos_antes = set(os.listdir(pasta))
    limite = time.time() + timeout

    while time.time() < limite:
        arquivos_agora = set(os.listdir(pasta))
        novos = [f for f in arquivos_agora - arquivos_antes if f.endswith(".csv")]
        if novos:
            arquivo_final = os.path.join(pasta, novos[0])
            print(f"[INFO] Novo arquivo detectado: {arquivo_final}")
            return arquivo_final
        time.sleep(1)

    raise TimeoutError("[ERRO] Nenhum novo arquivo .csv detectado após exportação.")

def realizar_login():
    print("[INFO] Acessando página de login...")
    driver.get("https://portal.soawebservices.com.br/Negativacoes/VisaoGeral")

    try:
        campo_email = wait.until(EC.visibility_of_element_located((By.ID, "Email")))
        campo_email.send_keys(LOGIN)
        print("[INFO] Campo de e-mail preenchido.")

        campo_senha = wait.until(EC.visibility_of_element_located((By.ID, "Senha")))
        campo_senha.send_keys(SENHA)
        print("[INFO] Campo de senha preenchido.")

        botao_login_seguro = wait.until(EC.element_to_be_clickable((By.ID, "js-login-btn")))
        time.sleep(1)
        botao_login_seguro.click()
        print("[INFO] Botão 'LoginSeguro' clicado.")

        wait.until(EC.presence_of_element_located((By.ID, "btn_Responsaveis")))
        print("[INFO] Login realizado com sucesso.")

    except Exception as e:
        print("[ERRO] Falha no processo de login:", e)
        with open("pagina_login_erro.html", "w", encoding="utf-8") as f:
            f.write(driver.page_source)
        print("[DEBUG] HTML salvo como 'pagina_login_erro.html'.")
        raise

def clicar_exportar_csv(botao_id, nome_saida):
    print(f"[INFO] Acessando aba '{botao_id}' com duplo clique...")
    try:
        # Tratar aba "Erros" sem ID específico
        if botao_id == "href_Erros":
            aba = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, 'a[href="#tab_Erros"]')))
        else:
            aba = wait.until(EC.element_to_be_clickable((By.ID, botao_id)))
        aba.click()
        time.sleep(1)
        aba.click()
        print(f"[INFO] Aba '{botao_id}' clicada duas vezes.")

        print("[INFO] Aguardando carregamento da aba...")
        time.sleep(2)

        print("[INFO] Procurando botão 'Exportar em CSV'...")

        # Mapeamento dos botões de exportação por aba
        if botao_id == "btn_Financeiro":
            botao_exportar = wait.until(EC.element_to_be_clickable((By.ID, "GridNegativacoesBaixadas_csvexport")))
        elif botao_id == "btn_Cobranca":
            botao_exportar = wait.until(EC.element_to_be_clickable((By.ID, "GridNegativacoesPendentes_csvexport")))
        elif botao_id == "btn_NFSe" and nome_saida == "Determinacao.csv":
            botao_exportar = wait.until(EC.element_to_be_clickable((By.ID, "GridNegativacoesRecusadas_csvexport")))
        elif botao_id == "href_Erros":
            botao_exportar = wait.until(EC.element_to_be_clickable((By.ID, "GridNegativacoesErros_csvexport")))
        else:
            botao_exportar = wait.until(EC.element_to_be_clickable((
                By.XPATH,
                "//span[text()='Exportar em CSV']/ancestor::button | //span[text()='Exportar em CSV']/ancestor::*[contains(@class, 'e-tbar-btn')]"
            )))

        driver.execute_script("arguments[0].style.border='2px solid red'", botao_exportar)

        try:
            botao_exportar.click()
        except Exception:
            driver.execute_script("arguments[0].click();", botao_exportar)

        print("[INFO] Botão 'Exportar em CSV' clicado.")

        arquivo = aguardar_download(download_dir, timeout=120)
        destino = os.path.join(download_dir, nome_saida)
        os.rename(arquivo, destino)
        print(f"[INFO] Arquivo renomeado para: {destino}")

    except Exception as e:
        print(f"[ERRO] Falha ao exportar aba '{botao_id}':", e)
        with open(f"pagina_{botao_id}_erro.html", "w", encoding="utf-8") as f:
            f.write(driver.page_source)
        print(f"[DEBUG] HTML salvo como 'pagina_{botao_id}_erro.html'.")
        raise

# ===== Execução principal =====

try:
    realizar_login()

    abas = [
        ("btn_Responsaveis", "Ativas.csv"),
        ("btn_Financeiro", "Baixadas.csv"),
        ("btn_Cobranca", "Pendentes.csv"),
        ("btn_NFSe", "Determinacao.csv"),
        ("href_Erros", "Erros.csv")
    ]

    for aba_id, nome_saida in abas:
        clicar_exportar_csv(aba_id, nome_saida)

except Exception as e:
    print(f"[ERRO GERAL] {e}")

finally:
    driver.quit()
    print("[INFO] Processo concluído.")
