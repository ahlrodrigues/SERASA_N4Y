# -*- coding: utf-8 -*-
"""
Pipeline XLSX-only para unificar CNM + SGP + SOA
- Mantém CNM em XLSX
- 'Tipo' sempre 'INCLUSÃO' ou 'EXCLUSÃO' (com acento)
- Saída final: output/dashboard_unificado.xlsx com colunas:
  Documento, Nome, SOA_Status, SGP_Status, CNM_Status, Operação, Tipo, Consolidado
"""

import os, re, glob, shutil, zipfile, platform, subprocess, urllib.request
from pathlib import Path
import pandas as pd
import unicodedata

# =========================
# ======== CONFIG =========
# =========================
DOWNLOAD_DIR = "./download"
OUTPUT_DIR   = "./output"

CNM_XLSX = os.path.join(DOWNLOAD_DIR, "Relatorio_CNM.xlsx")   # baixado via Selenium
SGP_XLSX = os.path.join(DOWNLOAD_DIR, "Relatorio_SGP.xlsx")

SOA_CSVS = {
    "Ativas":       os.path.join(DOWNLOAD_DIR, "Ativas.csv"),
    "Baixadas":     os.path.join(DOWNLOAD_DIR, "Baixadas.csv"),
    "Determinacao": os.path.join(DOWNLOAD_DIR, "Determinacao.csv"),
    "Erros":        os.path.join(DOWNLOAD_DIR, "Erros.csv"),
    "Pendentes":    os.path.join(DOWNLOAD_DIR, "Pendentes.csv"),
}

OUT_XLSX  = os.path.join(OUTPUT_DIR, "dashboard_unificado.xlsx")

# =========================
# ======= UTILS ===========
# =========================
def info(m): print(f"[INFO] {m}")
def warn(m): print(f"[WARN] {m}")

def _norm_text(s):
    if pd.isna(s): return ""
    s = str(s).strip()
    s = unicodedata.normalize("NFKD", s).encode("ascii", "ignore").decode("ascii")
    return s

def _norm_cols(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df.columns = [_norm_text(c).lower().replace(" ", "_") for c in df.columns]
    return df

def _as_digits(s): return re.sub(r"\D", "", str(s) if s is not None else "")

def _detect_header_row(df_raw: pd.DataFrame, max_scan=30):
    patterns = ["documento","cpf","cnpj","cpf_cnpj","nome","razao","fantasia"]
    for i in range(min(max_scan, len(df_raw))):
        row = df_raw.iloc[i].astype(str).str.lower().fillna("")
        if any(p in " ".join(row.tolist()) for p in patterns):
            return i
    return 0

def ensure_output_dir_clean():
    Path(OUTPUT_DIR).mkdir(parents=True, exist_ok=True)
    for p in Path(OUTPUT_DIR).glob("*"):
        if p.is_file(): p.unlink()
        elif p.is_dir(): shutil.rmtree(p)

# =========================
# ======= MAPEAMENTOS =====
# =========================
def cnm_tipo_para_status(tipo_raw: str) -> str:
    t = _norm_text(tipo_raw).upper()
    if t == "INCLUSAO":   return "NEGATIVADO"
    if t == "EXCLUSAO":   return "BAIXADO"
    if t == "PROCESSANDO": return ""  # ignorado na consolidação
    return "ERRO" if t else ""

def cnm_tipo_acento(tipo_raw: str) -> str:
    """Retorna 'INCLUSÃO' ou 'EXCLUSÃO' (com acento); caso contrário, '---'."""
    t = _norm_text(tipo_raw).upper()
    if t == "INCLUSAO": return "INCLUSÃO"
    if t == "EXCLUSAO": return "EXCLUSÃO"
    return "---"

# =========================
# ========= CNM ===========
# =========================
def ler_cnm_xlsx(path_xlsx: str) -> pd.DataFrame:
    if not os.path.exists(path_xlsx):
        raise FileNotFoundError(f"CNM não encontrado: {path_xlsx}")

    df_raw = pd.read_excel(path_xlsx, header=None)
    header_row = _detect_header_row(df_raw, max_scan=40)
    df_raw.columns = df_raw.iloc[header_row].astype(str).tolist()
    df_raw = df_raw.iloc[header_row+1:].reset_index(drop=True)
    df = _norm_cols(df_raw)

    aliases = {
        "documento": ["documento","cpf_cnpj","cpfcnpj","cpf","cnpj","doc","unnamed:_3","unnamed:_2"],
        "nome":      ["nome","nome_razao_social","razao_social","nome_fantasia","cliente","pessoa"],
        "operacao":  ["operacao","operacao_","tipo_operacao"],
        "tipo":      ["tipo","tipo_","status","situacao"]
    }

    def pick(keys):
        for k in keys:
            if k in df.columns: return k
        return None

    col_doc  = pick(aliases["documento"])
    col_nome = pick(aliases["nome"])  or "nome"
    col_op   = pick(aliases["operacao"]) or "operacao"
    col_tipo = pick(aliases["tipo"]) or "tipo"

    # Heurística p/ Documento se não achar
    if col_doc is None:
        best_col, best_ratio = None, -1
        for c in df.columns:
            series = df[c].astype(str).fillna("")
            digits = series.map(_as_digits)
            long_num = digits.map(lambda x: len(x) >= 9)
            ratio = long_num.mean()
            if ratio > best_ratio: best_ratio, best_col = ratio, c
        col_doc = best_col
        info(f"[CNM] Coluna 'documento' detectada: {col_doc} (ratio={best_ratio:.2f})")

    out = pd.DataFrame({
        "Documento": df[col_doc].astype(str).map(_as_digits),        # sem zfill para aceitar CNPJ
        "Nome":      df[col_nome].astype(str).str.strip(),
        "Operação":  df[col_op].astype(str).str.strip() if col_op in df.columns else "",
        "Tipo":      df[col_tipo].apply(cnm_tipo_acento)              # INCLUSÃO / EXCLUSÃO / ---
    })
    out["CNM_Status"] = df[col_tipo].apply(cnm_tipo_para_status)
    return out

# =========================
# ========= SGP ===========
# =========================
def ler_sgp_xlsx(path_xlsx: str) -> pd.DataFrame:
    if not os.path.exists(path_xlsx):
        warn(f"SGP não encontrado: {path_xlsx}. Voltando vazio.")
        return pd.DataFrame(columns=["Documento","SGP_Status"])

    df = pd.read_excel(path_xlsx)
    df = _norm_cols(df)

    # Documento
    doc_col = None
    for c in ["cpf_cnpj","cpfcnpj","cpf","cnpj","documento","doc"]:
        if c in df.columns: doc_col = c; break
    if doc_col is None:
        best_col, best_ratio = None, -1
        for c in df.columns:
            series = df[c].astype(str).fillna("")
            digits = series.map(_as_digits)
            long_num = digits.map(lambda x: len(x) >= 9)
            ratio = long_num.mean()
            if ratio > best_ratio: best_ratio, best_col = ratio, c
        doc_col = best_col
        info(f"[SGP] Coluna doc detectada: {doc_col} (ratio={best_ratio:.2f})")

    # Status
    status_col = None
    for c in ["status","situacao","tipo","resultado"]:
        if c in df.columns: status_col = c; break

    out = pd.DataFrame({"Documento": df[doc_col].astype(str).map(_as_digits)})
    if status_col:
        s = df[status_col].astype(str).str.upper().str.strip()
        s = s.replace({"INCLUSAO":"NEGATIVADO","EXCLUSAO":"BAIXADO"})
        out["SGP_Status"] = s.where(s.isin(["NEGATIVADO","BAIXADO"]), "ERRO")
    else:
        out["SGP_Status"] = ""
    # 1 linha por documento (pega primeiro válido)
    out = out.groupby("Documento", as_index=False)["SGP_Status"].agg(lambda x: next((v for v in x if v), ""))
    return out

# =========================
# ========= SOA ===========
# =========================
def ler_soa_csvs(csvs: dict) -> pd.DataFrame:
    frames = []
    for aba, path in csvs.items():
        if not os.path.exists(path): continue
        try:
            df = pd.read_csv(path, dtype=str, engine="python")
        except Exception:
            df = pd.read_csv(path, dtype=str, engine="python", sep=";")
        df = _norm_cols(df)

        # Documento
        doc_col = None
        for c in ["documento","cpf_cnpj","cpfcnpj","cpf","cnpj","doc"]:
            if c in df.columns: doc_col = c; break
        if doc_col is None:
            best_col, best_ratio = None, -1
            for c in df.columns:
                series = df[c].astype(str).fillna("")
                digits = series.map(_as_digits)
                long_num = digits.map(lambda x: len(x) >= 9)
                ratio = long_num.mean()
                if ratio > best_ratio: best_ratio, best_col = ratio, c
            doc_col = best_col
            info(f"[SOA/{aba}] Coluna doc detectada: {doc_col} (ratio={best_ratio:.2f})")

        tmp = pd.DataFrame({"Documento": df[doc_col].astype(str).map(_as_digits)})
        aba_l = aba.lower()
        if aba_l == "ativas":
            tmp["SOA_Status"] = "NEGATIVADO"; tmp["SOA_Origem"] = "Ativas"
        elif aba_l == "baixadas":
            tmp["SOA_Status"] = "BAIXADO";    tmp["SOA_Origem"] = "Baixadas"
        elif aba_l in ("determinacao","determinação"):
            tmp["SOA_Status"] = "ERRO";       tmp["SOA_Origem"] = "Determinacao"
        elif aba_l == "erros":
            tmp["SOA_Status"] = "ERRO";       tmp["SOA_Origem"] = "Erros"
        elif aba_l == "pendentes":
            tmp["SOA_Status"] = "";           tmp["SOA_Origem"] = "Pendentes"
        else:
            tmp["SOA_Status"] = "";           tmp["SOA_Origem"] = aba
        frames.append(tmp)

    if not frames:
        warn("Sem CSVs do SOA. Voltando vazio.")
        return pd.DataFrame(columns=["Documento","SOA_Status","SOA_Origem"])

    df_soa = pd.concat(frames, ignore_index=True)
    def pick_status(vals):
        vals = list(pd.Series(vals).dropna().astype(str))
        for pref in ["NEGATIVADO","BAIXADO","ERRO"]:
            if pref in vals: return pref
        return ""
    out = df_soa.groupby("Documento", as_index=False).agg({
        "SOA_Status": pick_status,
        "SOA_Origem": lambda x: next((v for v in x if v), "---")
    })
    return out

# =========================
# ===== CONSOLIDAÇÃO ======
# =========================
def consolidar_row(row):
    cnm_status = str(row.get("CNM_Status","")).upper().strip()
    soa_status = str(row.get("SOA_Status","")).upper().strip()
    sgp_status = str(row.get("SGP_Status","")).upper().strip()

    cnm_origem = str(row.get("CNM_Origem","")).strip() or "---"
    soa_origem = str(row.get("SOA_Origem","")).strip() or "---"

    # Regra especial solicitada
    if sgp_status == "NEGATIVADO" and cnm_origem == "---" and soa_origem == "---":
        return "ERRO"

    statuses = {cnm_status, soa_status, sgp_status} - {""}
    if not statuses: return "---"
    if "ERRO" in statuses: return "ERRO"
    if {"NEGATIVADO","BAIXADO"}.issubset(statuses): return "ERRO"
    if len(statuses) == 1: return list(statuses)[0]
    return "ERRO"

def unificar_cnm_soa_sgp(df_cnm, df_soa, df_sgp):
    base = df_cnm.merge(df_soa, on="Documento", how="outer")
    base = base.merge(df_sgp, on="Documento", how="outer")

    # Origens para a regra especial
    base["CNM_Origem"] = base["Operação"].where(base["Operação"].notna(), "---").fillna("---")
    base["SOA_Origem"] = base.get("SOA_Origem","---")
    base["SOA_Origem"] = base["SOA_Origem"].fillna("---")

    # Consolidado
    base["Consolidado"] = base.apply(consolidar_row, axis=1)

    # Remover PROCESSANDO do CNM (CNM_Status vazio) se nenhuma outra origem trouxe status
    mask_keep = (base["CNM_Status"].fillna("")!="") | (base["SOA_Status"].fillna("")!="") | (base["SGP_Status"].fillna("")!="")
    base = base[mask_keep].copy()

    # Colunas finais na ordem pedida
    finais = ["Documento","Nome","SOA_Status","SGP_Status","CNM_Status","Operação","Tipo","Consolidado"]
    # Preenche ausentes
    for c in finais:
        if c not in base.columns: base[c] = "---"
        base[c] = base[c].fillna("---").replace("", "---")

    return base[finais].sort_values(by=["Documento","Nome"]).reset_index(drop=True)

# =========================
# ========= MAIN ==========
# =========================
def gerar_xlsx_unificado():
    info("Lendo CNM (XLSX) preservando Operação e criando Tipo (INCLUSÃO/EXCLUSÃO)...")
    df_cnm = ler_cnm_xlsx(CNM_XLSX)

    info("Lendo SGP (XLSX)...")
    df_sgp = ler_sgp_xlsx(SGP_XLSX)

    info("Lendo SOA (CSVs)...")
    df_soa = ler_soa_csvs(SOA_CSVS)

    info("Construindo unificado...")
    df_final = unificar_cnm_soa_sgp(df_cnm, df_soa, df_sgp)

    ensure_output_dir_clean()
    info(f"Gravando XLSX final com {len(df_final)} linhas em: {OUT_XLSX}")
    with pd.ExcelWriter(OUT_XLSX, engine="xlsxwriter") as writer:
        df_final.to_excel(writer, index=False, sheet_name="Dashboard")
        # Larguras básicas
        ws = writer.sheets["Dashboard"]
        widths = {
            "Documento":18, "Nome":36, "SOA_Status":14, "SGP_Status":14, "CNM_Status":14,
            "Operação":18, "Tipo":12, "Consolidado":14
        }
        for idx, col in enumerate(df_final.columns, start=0):
            ws.set_column(idx, idx, widths.get(col, 16))

    info("Concluído.")

if __name__ == "__main__":
    gerar_xlsx_unificado()
