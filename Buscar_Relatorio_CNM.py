# Buscar_Relatorio_CNM.py
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
from dotenv import load_dotenv

import os
import time
import subprocess
import zipfile
import urllib.request
import shutil
import platform
import re
from pathlib import Path

# ==== Pós-processamento ====
import pandas as pd
import unicodedata
from openpyxl import load_workbook

# =========================
# ==== Funções util =======
# =========================
def _norm_text(s):
    if pd.isna(s):
        return ""
    s = str(s).strip()
    s = unicodedata.normalize("NFKD", s).encode("ascii", "ignore").decode("ascii")
    return s

def _norm_cols(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df.columns = [_norm_text(c).lower().replace(" ", "_") for c in df.columns]
    return df

def _as_digits(s):
    return re.sub(r"\D", "", str(s) if s is not None else "")

def _detect_header_row(df_raw: pd.DataFrame, max_scan=40):
    patterns = ["documento","cpf","cnpj","cpf_cnpj","nome","razao","fantasia"]
    for i in range(min(max_scan, len(df_raw))):
        row = df_raw.iloc[i].astype(str).str.lower().fillna("")
        if any(p in " ".join(row.tolist()) for p in patterns):
            return i
    return 0

def status_por_tipo_cnm(tipo: str) -> str:
    t = _norm_text(tipo).upper()
    if t == "INCLUSAO":    return "NEGATIVADO"
    if t == "EXCLUSAO":    return "BAIXADO"
    if t == "PROCESSANDO": return ""  # ignorar no dash
    return "ERRO" if t else ""

def tipo_cnm_com_acento(tipo: str) -> str:
    t = _norm_text(tipo).upper()
    if t == "INCLUSAO": return "INCLUSÃO"
    if t == "EXCLUSAO": return "EXCLUSÃO"
    return "---"

def is_lock_or_temp(name: str) -> bool:
    name = name.lower()
    if name.endswith(".crdownload"):
        return True
    if name.startswith(".~lock.") and name.endswith(".xlsx#"):
        return True
    if name.endswith((".tmp", ".part", ".partial", "~")):
        return True
    return False

def is_valid_xlsx(path: str) -> bool:
    try:
        if not os.path.exists(path):
            return False
        if os.path.getsize(path) < 4096:
            return False
        if not zipfile.is_zipfile(path):
            return False
        wb = load_workbook(filename=path, read_only=True)
        wb.close()
        return True
    except Exception:
        return False

def sanitize_filename(name: str) -> str:
    name = re.sub(r'[:*?"<>|\\/\r\n\t]', "_", str(name)).strip().strip(".")
    return name or "arquivo"

# =========================
# ==== Pós-processar CNM ==
# =========================
def posprocessar_cnm(caminho_xlsx: str, saida_xlsx_standard: str):
    """
    Lê o XLSX do CNM, detecta o cabeçalho e salva XLSX padronizado:
    Documento, Nome, Operação, Tipo (INCLUSÃO/EXCLUSÃO), CNM_Status
    """
    print("🛠️ Pós-processando CNM para manter 'Operação' e 'Tipo'...")

    if not is_valid_xlsx(caminho_xlsx):
        raise ValueError(f"Arquivo inválido para pós-processamento: {caminho_xlsx}")

    df_raw = pd.read_excel(caminho_xlsx, header=None, engine="openpyxl")
    header_row = _detect_header_row(df_raw, max_scan=50)
    df_raw.columns = df_raw.iloc[header_row].astype(str).tolist()
    df_raw = df_raw.iloc[header_row+1:].reset_index(drop=True)

    df = _norm_cols(df_raw)

    aliases = {
        "documento": ["documento","cpf_cnpj","cpfcnpj","cpf","cnpj","doc","unnamed:_3","unnamed:_2"],
        "nome":      ["nome","nome_razao_social","razao_social","nome_fantasia","cliente","pessoa"],
        "operacao":  ["operacao","operacao_","tipo_operacao"],
        "tipo":      ["tipo","tipo_","status","situacao"],
    }

    def pick(keys):
        for k in keys:
            if k in df.columns: return k
        return None

    col_doc  = pick(aliases["documento"])
    col_nome = pick(aliases["nome"])  or "nome"
    col_op   = pick(aliases["operacao"]) or "operacao"
    col_tipo = pick(aliases["tipo"]) or "tipo"

    if col_doc is None:
        best_col, best_ratio = None, -1
        for c in df.columns:
            r = df[c].astype(str).map(lambda v: len(_as_digits(v))>=9).mean()
            if r > best_ratio: best_ratio, best_col = r, c
        col_doc = best_col
        print(f"ℹ️ Coluna 'documento' detectada por amostragem: {col_doc} (ratio={best_ratio:.2f})")

    out = pd.DataFrame({
        "Documento": df[col_doc].astype(str).map(_as_digits),  # aceita CPF e CNPJ
        "Nome":      df[col_nome].astype(str).str.strip(),
        "Operação":  df[col_op].astype(str).str.strip(),
        "Tipo":      df[col_tipo].apply(tipo_cnm_com_acento),
    })
    out["CNM_Status"] = df[col_tipo].map(status_por_tipo_cnm)

    os.makedirs(os.path.dirname(saida_xlsx_standard), exist_ok=True)
    saida_path = Path(saida_xlsx_standard)
    with pd.ExcelWriter(saida_path, engine="openpyxl") as writer:
        out.to_excel(writer, index=False, sheet_name="CNM")
    print(f"✅ CNM padrão salvo em XLSX: {saida_path}")

    faltantes = [c for c in ["Operação","Tipo"] if c not in out.columns]
    if faltantes:
        print(f"⚠️ Colunas faltantes no padrão: {faltantes}")
    else:
        print("✅ Colunas 'Operação' e 'Tipo' preservadas com sucesso.")

# =========================
# === ChromeDriver auto ===
# =========================
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
        versao_completa = resultado.stdout.strip().split()[-1]
        versao_principal = versao_completa.split(".")[0]
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

# =========================
# ====== Selenium =========
# =========================
def aguardar_download(pasta, timeout=180, quiet_stable_secs=2):
    """
    Espera surgir um XLSX **válido**:
    - ignora .crdownload e .~lock...xlsx#
    - exige extensão .xlsx
    - exige tamanho estável por quiet_stable_secs
    - valida ZIP Excel
    """
    print("⏳ Aguardando o download finalizar (XLSX válido)...")
    limite = time.time() + timeout
    ultimo_tamanho = {}
    ultimo_momento = {}

    while time.time() < limite:
        for nome in os.listdir(pasta):
            if is_lock_or_temp(nome):
                continue
            if not nome.lower().endswith(".xlsx"):
                continue

            full = os.path.join(pasta, nome)
            tam = os.path.getsize(full)
            prev = ultimo_tamanho.get(full)
            now  = time.time()

            if prev is None or prev != tam:
                ultimo_tamanho[full] = tam
                ultimo_momento[full] = now
                continue

            if now - ultimo_momento[full] < quiet_stable_secs:
                continue

            if is_valid_xlsx(full):
                print(f"✅ Download completo e válido: {full}")
                return full
            else:
                print(f"⚠️ Arquivo .xlsx inválido detectado (ignorando): {full}")

        time.sleep(1)

    raise TimeoutError("❌ Tempo limite ao esperar um XLSX válido.")

# =========================
# ======== MAIN ===========
# =========================
if __name__ == "__main__":
    # Config inicial
    driver_path = "./chromedriver"
    instalar_chromedriver_compatível(driver_path)

    load_dotenv()
    USUARIO = os.getenv("USUARIO_CNM")
    SENHA   = os.getenv("SENHA_CNM")

    if not USUARIO or not SENHA:
        raise RuntimeError("Defina USUARIO_CNM e SENHA_CNM no .env")

    diretorio_download = os.path.abspath("download")
    os.makedirs(diretorio_download, exist_ok=True)

    # Chrome
    chrome_options = Options()
    chrome_options.add_experimental_option("prefs", {
        "download.default_directory": diretorio_download,
        "download.prompt_for_download": False,
        "download.directory_upgrade": True,
        "safebrowsing.enabled": True
    })

    # Navegador
    driver = webdriver.Chrome(service=Service(driver_path), options=chrome_options)
    driver.maximize_window()
    wait = WebDriverWait(driver, 20)

    # Fluxo CNM
    driver.get("https://appv2.creditonamedida.com.br/logar")

    campo_usuario = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, 'input[placeholder="Digite seu usuário"]')))
    campo_usuario.clear(); campo_usuario.send_keys(USUARIO)

    campo_senha = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, 'input[placeholder="Digite sua senha"]')))
    campo_senha.clear(); campo_senha.send_keys(SENHA)

    botao_entrar = wait.until(EC.element_to_be_clickable((By.XPATH, '//button[.//span[text()="Logar"]]')))
    botao_entrar.click()

    wait.until(EC.url_changes("https://appv2.creditonamedida.com.br/logar"))
    print("✅ Login realizado com sucesso.")

    print("➡️ Localizando o botão 'Relatórios'...")
    menu_relatorios = wait.until(EC.element_to_be_clickable((By.XPATH, '//a[contains(text(), "Relatórios")]')))
    menu_relatorios.click(); time.sleep(1); menu_relatorios.click()
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
    campo_data_inicial.clear(); campo_data_inicial.send_keys("01012000")
    print("📅 Data inicial preenchida.")

    botao_pesquisar = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, 'button[title="Pesquisar"]')))
    botao_pesquisar.click()
    print("🔍 Botão 'Pesquisar' clicado.")
    time.sleep(5)

    # Remove antigo
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
        arquivo_original = aguardar_download(diretorio_download, timeout=180)
        novo_nome = relatorio_cnm
        os.rename(arquivo_original, novo_nome)
        print("✅ Arquivo renomeado para:", novo_nome)

        # Valida novamente pós-rename
        if not is_valid_xlsx(novo_nome):
            raise ValueError(f"Arquivo renomeado não é um XLSX válido: {novo_nome}")
    except Exception as e:
        print("❌ Erro no download/validação do XLSX:", e)
        driver.quit()
        raise

    driver.quit()
    print("✅ Processo de download concluído.")

    # Pós-processamento para XLSX standard
    standard_xlsx = os.path.join(diretorio_download, "Relatorio_CNM_standard.xlsx")
    try:
        posprocessar_cnm(relatorio_cnm, standard_xlsx)
    except Exception as e:
        print("❌ Falha no pós-processamento CNM:", e)
