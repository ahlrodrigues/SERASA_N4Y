# -*- coding: utf-8 -*-
import os
import time
import subprocess
import zipfile
import urllib.request
import shutil
from pathlib import Path
from dotenv import load_dotenv, dotenv_values

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

# ==============================
# Config fixas (URLs)
# ==============================
VISAO_URL = "https://portal.soawebservices.com.br/Negativacoes/VisaoGeral"
POS_LOGIN_SINAL = (
    By.XPATH,
    "//span[normalize-space()='Exportar em CSV']/ancestor::*[contains(@class,'e-tbar-btn') or self::button]"
)

# ==============================
# Atualizador automático do ChromeDriver
# ==============================
def instalar_chromedriver_compatível(driver_path="./chromedriver"):
    print("🔍 Verificando versão do Google Chrome instalada...")
    try:
        resultado = subprocess.run(
            ["/opt/google/chrome/chrome", "--version"],
            stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True, check=True
        )
        versao_completa = resultado.stdout.strip().split()[-1]
        print(f"✅ Chrome instalado: versão {versao_completa}")
    except Exception as e:
        print("❌ Não foi possível obter a versão do Google Chrome:", e)
        return False

    print("🌐 Buscando ChromeDriver compatível...")
    try:
        url_zip = f"https://storage.googleapis.com/chrome-for-testing-public/{versao_completa}/linux64/chromedriver-linux64.zip"
        caminho_zip = "chromedriver.zip"

        urllib.request.urlretrieve(url_zip, caminho_zip)
        print("📦 Download do ChromeDriver concluído.")

        with zipfile.ZipFile(caminho_zip, 'r') as zip_ref:
            zip_ref.extractall("chromedriver_temp")

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
        print("❌ Falha ao baixar/substituir ChromeDriver:", e)
        return False

# ==============================
# Setup inicial
# ==============================
driver_path = "./chromedriver"
instalar_chromedriver_compatível(driver_path)

dotenv_path = os.path.join(os.path.dirname(__file__), ".env")
load_dotenv(dotenv_path)
config = dotenv_values(dotenv_path)
print("[DEBUG] .env:", config)

LOGIN = config.get("USUARIO_SOA")
SENHA = config.get("SENHA_SOA")

if not LOGIN or not SENHA:
    print("[ERRO] USUARIO_SOA ou SENHA_SOA ausentes/vazios no .env")
    print(f"[DEBUG] USUARIO_SOA = {LOGIN}")
    print(f"[DEBUG] SENHA_SOA = {'<vazio>' if not SENHA else '***'}")
    raise SystemExit(1)

# Pasta de download + limpeza
download_dir = os.path.abspath("download")
Path(download_dir).mkdir(parents=True, exist_ok=True)
print(f"📁 Pasta de download pronta: {download_dir}")

for nome in ["Ativas.csv", "Baixadas.csv", "Determinacao.csv", "Erros.csv", "Pendentes.csv"]:
    alvo = os.path.join(download_dir, nome)
    if os.path.exists(alvo):
        try:
            os.remove(alvo)
            print(f"🪚 Removido: {nome}")
        except Exception as e:
            print(f"⚠️ Não foi possível remover {nome}: {e}")

# ==============================
# Chrome Options (corrigidas)
# ==============================
chrome_options = Options()
prefs = {
    "download.default_directory": download_dir,
    "download.prompt_for_download": False,
    "download.directory_upgrade": True,  # CORRETO
    "safebrowsing.enabled": True,
}
chrome_options.add_experimental_option("prefs", prefs)
# chrome_options.add_argument("--headless=new")  # descomente se quiser headless
chrome_options.add_argument("--no-sandbox")
chrome_options.add_argument("--disable-dev-shm-usage")
chrome_options.add_argument("--window-size=1366,768")
chrome_options.add_argument("--lang=pt-BR")
chrome_options.add_argument("--disable-gpu")
chrome_options.add_argument("--remote-allow-origins=*")

# Cria driver com logs úteis
service = Service(driver_path)
driver = webdriver.Chrome(service=service, options=chrome_options)
print("[INFO] Chrome aberto.")
print("[DEBUG] Versões:", driver.capabilities.get("browserVersion"), "|",
      driver.capabilities.get("chrome", {}).get("chromedriverVersion"))

wait = WebDriverWait(driver, 20)

# ==============================
# Helpers
# ==============================
def aguardar_download(pasta, timeout=180):
    """Espera por novo .csv. Simples e confiável para grids com export síncrono."""
    print("[INFO] Aguardando novo arquivo .csv na pasta de download...")
    inicio = time.time()
    baseline = set(os.listdir(pasta))
    while time.time() - inicio < timeout:
        atuais = set(os.listdir(pasta))
        novos = [f for f in atuais - baseline if f.lower().endswith(".csv")]
        if novos:
            caminho = os.path.join(pasta, sorted(novos)[-1])
            print(f"[INFO] Novo arquivo: {caminho}")
            return caminho
        time.sleep(1)
    raise TimeoutError("[ERRO] Nenhum .csv detectado após exportação.")

def esperar_overlays_sumirem(timeout=10):
    """Tenta aguardar overlays/spinners comuns sumirem, para evitar click intercepted."""
    try:
        WebDriverWait(driver, timeout).until(
            EC.invisibility_of_element_located((By.CSS_SELECTOR, ".e-overlay, .e-spinner-pane, .blockUI, .loading, .modal-backdrop"))
        )
    except Exception:
        pass

# ==============================
# Login (via Visão Geral)
# ==============================
def realizar_login():
    print("[INFO] Acessando Visão Geral...")
    driver.get(VISAO_URL)

    try:
        # Espera por UM dos dois: (a) form de login, (b) sinal de página pós-login
        def any_ready(drv):
            try:
                if drv.find_elements(By.ID, "Email") and drv.find_elements(By.ID, "Senha"):
                    return "LOGIN"  # formulário presente
                if drv.find_elements(*POS_LOGIN_SINAL):
                    return "OK"     # já logado na Visão Geral
            except Exception:
                pass
            return False

        estado = WebDriverWait(driver, 20).until(any_ready)
        print(f"[DEBUG] Estado inicial após abrir Visão Geral: {estado}")

        if estado == "OK":
            print("[INFO] Já logado. Seguindo...")
            return

        # === Fluxo de login (form exibido) ===
        campo_email = wait.until(EC.visibility_of_element_located((By.ID, "Email")))
        campo_email.clear(); campo_email.send_keys(LOGIN)
        print("[INFO] E-mail preenchido.")

        campo_senha = wait.until(EC.visibility_of_element_located((By.ID, "Senha")))
        campo_senha.clear(); campo_senha.send_keys(SENHA)
        print("[INFO] Senha preenchida.")

        # Tente o botão 'LoginSeguro' por ID; se não houver, tente um submit do form
        try:
            botao = wait.until(EC.element_to_be_clickable((By.ID, "js-login-btn")))
            time.sleep(0.3)
            botao.click()
            print("[INFO] Botão 'LoginSeguro' clicado.")
        except Exception:
            from selenium.webdriver.common.keys import Keys
            campo_senha.send_keys(Keys.ENTER)
            print("[INFO] Form submetido via ENTER (fallback).")

        # Aguarda cair na Visão Geral (ou qualquer elemento que prove login)
        wait.until(EC.presence_of_element_located(POS_LOGIN_SINAL))
        print("[INFO] Login concluído e Visão Geral carregada.")

    except Exception as e:
        print("[ERRO] Falha ao abrir/logar pela Visão Geral:", e)
        os.makedirs("output", exist_ok=True)
        try:
            driver.save_screenshot("output/visao_geral_login_error.png")
        except Exception:
            pass
        try:
            with open("output/visao_geral_login_error.html", "w", encoding="utf-8") as f:
                f.write(driver.page_source)
        except Exception:
            pass
        raise

# ==============================
# Export direto por grid (sem clicar abas)
# ==============================
def exportar_por_grid(export_id: str, nome_saida: str, fallback_tab_id: str | None = None):
    """
    Tenta clicar diretamente no botão de export do grid. Se não achar,
    abre a aba fallback (uma vez) e tenta de novo.
    """
    print(f"[INFO] Exportando: {nome_saida} via #{export_id}")

    def _tentar_click_export() -> bool:
        try:
            esperar_overlays_sumirem(8)
            btn = wait.until(EC.element_to_be_clickable((By.ID, export_id)))
            driver.execute_script("arguments[0].scrollIntoView({block:'center'});", btn)
            time.sleep(0.2)
            try:
                btn.click()
            except Exception:
                driver.execute_script("arguments[0].click();", btn)
            print("[INFO] 'Exportar em CSV' clicado.")
            return True
        except Exception:
            return False

    # 1) Tenta direto
    if not _tentar_click_export():
        print(f"[WARN] Botão #{export_id} não visível ainda.")
        # 2) Fallback: abrir a aba correspondente (se fornecida) e tentar de novo
        if fallback_tab_id:
            print(f"[INFO] Abrindo aba fallback '{fallback_tab_id}'...")
            try:
                if fallback_tab_id == "href_Erros":
                    aba = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, 'a[href=\"#tab_Erros\"]')))
                else:
                    aba = wait.until(EC.element_to_be_clickable((By.ID, fallback_tab_id)))
                driver.execute_script("arguments[0].scrollIntoView({block:'center'});", aba)
                time.sleep(0.2)
                aba.click()
                time.sleep(1.2)  # tempo para grid renderizar
            except Exception as e:
                print(f"[WARN] Não consegui abrir a aba '{fallback_tab_id}': {e}")

        if not _tentar_click_export():
            raise RuntimeError(f"Não foi possível clicar no export #{export_id} (mesmo após fallback).")

    # 3) Espera o download aparecer e renomeia
    arquivo = aguardar_download(download_dir, timeout=180)
    destino = os.path.join(download_dir, nome_saida)
    try:
        if os.path.exists(destino):
            os.remove(destino)
    except Exception:
        base, ext = os.path.splitext(destino)
        destino = f"{base}_{int(time.time())}{ext}"
    os.rename(arquivo, destino)
    print(f"[INFO] CSV salvo: {destino}")

# ==============================
# Execução principal
# ==============================
try:
    realizar_login()

    targets = [
        # export_id                            , arquivo            , aba_fallback (se precisar)
        ("GridNegativacoesAtivas_csvexport",     "Ativas.csv",        "btn_Responsaveis"),
        ("GridNegativacoesBaixadas_csvexport",   "Baixadas.csv",      "btn_Financeiro"),
        ("GridNegativacoesPendentes_csvexport",  "Pendentes.csv",     "btn_Cobranca"),
        ("GridNegativacoesRecusadas_csvexport",  "Determinacao.csv",  "btn_NFSe"),
        ("GridNegativacoesErros_csvexport",      "Erros.csv",         "href_Erros"),
    ]

    for export_id, nome_saida, fallback_tab in targets:
        exportar_por_grid(export_id, nome_saida, fallback_tab)

except Exception as e:
    print(f"[ERRO GERAL] {e}")
finally:
    try:
        driver.quit()
    except Exception:
        pass
    print("[INFO] Processo concluído.")
