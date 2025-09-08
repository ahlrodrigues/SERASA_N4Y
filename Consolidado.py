# Consolidado.py
import os, re, glob
from pathlib import Path
import numpy as np
import pandas as pd
import unicodedata

# ===== Writer engine com fallback =====
try:
    import xlsxwriter  # noqa: F401
    ENGINE_WRITE = "xlsxwriter"
except Exception:
    ENGINE_WRITE = "openpyxl"

# ===== Paths =====
BASE_DIR = Path(__file__).resolve().parent
DOWNLOAD_DIR = (BASE_DIR / "download").resolve()
OUTPUT_DIR   = (BASE_DIR / "output").resolve()
OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

CNM_STD_XLSX = DOWNLOAD_DIR / "Relatorio_CNM_standard.xlsx"
CNM_RAW_XLSX = DOWNLOAD_DIR / "Relatorio_CNM.xlsx"
SGP_XLSX     = DOWNLOAD_DIR / "Relatorio_SGP.xlsx"
SOA_CSVS     = {
    "Ativas":       DOWNLOAD_DIR / "Ativas.csv",
    "Baixadas":     DOWNLOAD_DIR / "Baixadas.csv",
    "Determinacao": DOWNLOAD_DIR / "Determinacao.csv",
    "Erros":        DOWNLOAD_DIR / "Erros.csv",
    "Pendentes":    DOWNLOAD_DIR / "Pendentes.csv",
}
OUT_XLSX = OUTPUT_DIR / "dashboard_unificado.xlsx"

# ===== Config SGP =====
SGP_HEADER_ROW = 8  # "linha 9" (0-based)

# ===== Logs =====
def info(m): print(f"[INFO] {m}")
def warn(m): print(f"[WARN] {m}")

# ===== Utils =====
def _norm_text(s):
    if pd.isna(s): return ""
    s = str(s).strip()
    s = unicodedata.normalize("NFKD", s).encode("ascii", "ignore").decode("ascii")
    return s

def _norm_cols(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    new_cols = []
    for c in df.columns:
        base = _norm_text(c).lower()
        base = re.sub(r"[^a-z0-9]+", "_", base).strip("_")
        new_cols.append(base)
    df.columns = new_cols
    return df

def _digits(s): return re.sub(r"\D","", str(s) if s is not None else "")

def _valid_doc_mask(series: pd.Series) -> pd.Series:
    return series.astype(str).map(_digits).str.len().isin([11, 14])

def _detect_header_row(df_raw: pd.DataFrame, max_scan=120):
    tokens = ["documento","cpf","cnpj","cpf_cnpj","nome","razao","fantasia",
              "status","situacao","resultado","operacao","tipo","acao","retorno","data","dt","ocorr"]
    best_i, best_score = 0, -1
    for i in range(min(max_scan, len(df_raw))):
        row = df_raw.iloc[i].astype(str).fillna("").tolist()
        rown = [_norm_text(x).lower() for x in row]
        nonempty = sum(1 for x in rown if x != "")
        hits = sum(1 for x in rown for t in tokens if t in x)
        score = hits*3 + nonempty
        if score > best_score:
            best_score, best_i = score, i
    return best_i

def _tipo_com_acento(tipo_raw: str) -> str:
    t = _norm_text(tipo_raw).upper()
    if t == "INCLUSAO": return "INCLUSÃO"
    if t == "EXCLUSAO": return "EXCLUSÃO"
    return "---"

def _status_por_tipo(tipo_raw: str) -> str:
    t = _norm_text(tipo_raw).upper()
    if t == "INCLUSAO":    return "NEGATIVADO"
    if t == "EXCLUSAO":    return "BAIXADO"
    if t == "PROCESSANDO": return ""
    return "ERRO" if t else ""

def pick_name_column(df: pd.DataFrame, exclude=()):
    import re as _re
    exclude = set(exclude or ())
    candidates = []
    for c in df.columns:
        if c in exclude:
            continue
        cl = str(c).lower()
        if any(tok in cl for tok in [
            "usuario","user","login","email","e_mail","mail",
            "operacao","tipo","status","situacao","situação",
            "data","hora","dt","cod","codigo","código","id",
            "documento","cpf","cnpj","doc","chave","hash","origem"
        ]):
            continue
        s = df[c].astype(str).fillna("")
        if (s == "").mean() > 0.90:
            continue
        letters = s.map(lambda v: 1 if _re.search(r"[A-Za-zÀ-ÿ]{3,}", v) else 0).mean()
        tokens  = s.map(lambda v: len([t for t in _re.split(r"\s+", v.strip()) if t])).mean()
        uniq    = s.nunique(dropna=True) / max(1, (s != "").sum())
        avglen  = s.map(len).mean()
        length_penalty = 1.0 if 5 <= avglen <= 60 else 0.6
        token_bonus = min(tokens/2, 1.0)
        score = letters * 0.6 + token_bonus * 0.2 + uniq * 0.2
        score *= length_penalty
        candidates.append((score, c))
    if not candidates: return None
    candidates.sort(reverse=True)
    return candidates[0][1]

def find_latest(patterns):
    paths = []
    for pat in patterns:
        paths.extend(glob.glob(str(DOWNLOAD_DIR / pat)))
    if not paths:
        return None
    paths = sorted(paths, key=lambda p: os.path.getmtime(p), reverse=True)
    return Path(paths[0])

# ---- Datas helpers ----
DATE_NAME_HINTS = ("data","dt","ocorr","emiss","inclus","exclus","baixa","registro","atualiz")

def _parse_dates_from_df(df: pd.DataFrame) -> dict:
    out = {}
    for c in df.columns:
        cl = str(c).lower()
        if any(h in cl for h in DATE_NAME_HINTS):
            s = pd.to_datetime(df[c], errors="coerce", dayfirst=True)
            if (~s.isna()).mean() >= 0.02:
                out[c] = s
    if not out:
        for c in df.columns:
            s = pd.to_datetime(df[c], errors="coerce", dayfirst=True)
            if (~s.isna()).mean() >= 0.10:
                out[c] = s
    return out

def _pick_best_date(series_map: dict) -> str | None:
    if not series_map: return None
    cov = {c: (~s.isna()).mean() for c, s in series_map.items()}
    return max(cov, key=cov.get)

def _coalesce_name(*vals) -> str:
    vals = [(_norm_text(v)) for v in vals if _norm_text(v)]
    if not vals: return ""
    return max(vals, key=len)

# ===== CNM =====
def ler_cnm() -> pd.DataFrame:
    if CNM_STD_XLSX.exists():
        info(f"Lendo CNM standard: {CNM_STD_XLSX}")
        df = pd.read_excel(CNM_STD_XLSX, engine="openpyxl")
        df["Documento"] = df["Documento"].astype(str).map(_digits)
        df = df[_valid_doc_mask(df["Documento"])].copy()
        df["Nome"]      = df.get("Nome","").astype(str).str.strip()
        df["Tipo"]      = df.get("Tipo","---").apply(_tipo_com_acento)
        if "CNM_Status" not in df.columns:
            df["CNM_Status"] = df["Tipo"].map(lambda x: _status_por_tipo("INCLUSAO" if x=="INCLUSÃO" else ("EXCLUSAO" if x=="EXCLUSÃO" else x)))
        agg = df.groupby("Documento", as_index=False).agg(
            CNM_Qtd=("Documento","size"),
            Nome=("Nome", lambda s: _coalesce_name(*s)),
            Tipo=("Tipo", lambda s: s.dropna().astype(str).replace("", np.nan).dropna().tail(1).values[0] if s.notna().any() else "---"),
            CNM_Status=("CNM_Status", lambda s: s.dropna().astype(str).replace("", np.nan).dropna().tail(1).values[0] if s.notna().any() else "")
        )
        agg["CNM_Data"] = pd.NaT
        return agg[["Documento","Nome","Tipo","CNM_Status","CNM_Qtd","CNM_Data"]]

    if not CNM_RAW_XLSX.exists():
        warn(f"CNM não encontrado: {CNM_STD_XLSX} nem {CNM_RAW_XLSX}")
        return pd.DataFrame(columns=["Documento","Nome","Tipo","CNM_Status","CNM_Qtd","CNM_Data"])

    info(f"Lendo CNM bruto: {CNM_RAW_XLSX}")
    df_raw = pd.read_excel(CNM_RAW_XLSX, header=None, engine="openpyxl")
    header_row = _detect_header_row(df_raw)
    df_raw.columns = df_raw.iloc[header_row].astype(str).tolist()
    df_raw = df_raw.iloc[header_row+1:].reset_index(drop=True)
    dfn = _norm_cols(df_raw)

    aliases = {
        "documento": ["documento","cpf_cnpj","cpfcnpj","cpf","cnpj","doc","unnamed_3","unnamed_2"],
        "nome":      ["nome","nome_razao_social","razao_social","nome_fantasia","cliente","pessoa","nome_razao","razao"],
        "tipo":      ["tipo","tipo_","status","situacao"]
    }
    def pick(keys):
        for k in keys:
            if k in dfn.columns: return k
        return None

    col_doc  = pick(aliases["documento"])
    col_tipo = pick(aliases["tipo"])
    col_nome = pick(aliases["nome"])

    if col_doc is None:
        best, best_r = None, -1
        for c in dfn.columns:
            r = dfn[c].astype(str).map(lambda v: len(_digits(v)) in (11,14)).mean()
            if r > best_r: best_r, best = r, c
        col_doc = best
        info(f"[CNM] doc detectado: {col_doc} (ratio={best_r:.2f})")

    if col_nome is None:
        col_nome = pick_name_column(dfn, exclude={col_doc, col_tipo})
        if col_nome:
            info(f"[CNM] 'nome' por heurística: {col_nome}")
        else:
            dfn["nome"] = ""
            col_nome = "nome"
            warn("[CNM] Nenhuma coluna de nome encontrada; criando 'nome' vazia.")

    if col_tipo is None:
        dfn["tipo"] = ""
        col_tipo = "tipo"

    date_map = _parse_dates_from_df(dfn)
    cnm_incl_col = next((c for c in date_map if "inclu" in c), None)
    cnm_excl_col = next((c for c in date_map if "exclu" in c or "baix" in c), None)
    best_date_col = _pick_best_date(date_map)

    tmp = pd.DataFrame({
        "Documento": dfn[col_doc].astype(str).map(_digits),
        "Nome":      dfn[col_nome].astype(str).str.strip(),
        "Tipo":      dfn[col_tipo].apply(_tipo_com_acento),
        "CNM_Status": dfn[col_tipo].map(_status_por_tipo),
    })
    cnm_date = None
    if cnm_incl_col or cnm_excl_col:
        s_in = date_map.get(cnm_incl_col, pd.Series(pd.NaT, index=dfn.index))
        s_ex = date_map.get(cnm_excl_col, pd.Series(pd.NaT, index=dfn.index))
        cnm_date = np.where(tmp["Tipo"].eq("INCLUSÃO"), s_in, s_ex)
        cnm_date = pd.to_datetime(cnm_date, errors="coerce")
    else:
        cnm_date = date_map.get(best_date_col, pd.Series(pd.NaT, index=dfn.index))
    tmp["CNM_Data"] = cnm_date
    tmp = tmp[_valid_doc_mask(tmp["Documento"])].copy()

    def agg_cnm(g):
        qtd = len(g)
        nome = _coalesce_name(*g["Nome"])
        if g["CNM_Data"].notna().any():
            ridx = g["CNM_Data"].idxmax()
            st = str(g.loc[ridx, "CNM_Status"])
            tp = str(g.loc[ridx, "Tipo"])
            dt = g.loc[ridx, "CNM_Data"]
        else:
            st = g["CNM_Status"].replace("", np.nan).dropna().tail(1).values[0] if (g["CNM_Status"]!="").any() else ""
            tp = g["Tipo"].replace("", np.nan).dropna().tail(1).values[0] if (g["Tipo"]!="").any() else "---"
            dt = pd.NaT
        return pd.Series({"CNM_Qtd": qtd, "Nome": nome, "CNM_Status": st, "Tipo": tp, "CNM_Data": dt})

    agg = tmp.groupby("Documento", as_index=False).apply(agg_cnm)
    return agg[["Documento","Nome","Tipo","CNM_Status","CNM_Qtd","CNM_Data"]]

# ===== SGP (presença + nome) =====
def _pick_doc_col(df: pd.DataFrame) -> str:
    best, best_score = None, -1
    for c in df.columns:
        s = df[c].astype(str).map(_digits)
        ratio = ((s.str.len()==11) | (s.str.len()==14)).mean()
        uniq  = s.nunique(dropna=True)/max(1, (s!="").sum())
        score = ratio*0.8 + uniq*0.2
        if score > best_score:
            best_score, best = score, c
    return best

def ler_sgp() -> pd.DataFrame:
    try:
        path = SGP_XLSX if SGP_XLSX.exists() else find_latest(
            ["*SGP*.xlsx", "*sgp*.xlsx", "*SGP*.xls", "*sgp*.xls", "*SGP*.csv", "*sgp*.csv"]
        )
        if not path or not Path(path).exists():
            warn(f"SGP não encontrado: {SGP_XLSX} (e nenhum *SGP*.xlsx/.xls/.csv em {DOWNLOAD_DIR})")
            return pd.DataFrame(columns=["Documento","SGP_Status","Nome_SGP"])
        info(f"Lendo SGP de: {path}")

        if str(path).lower().endswith(".csv"):
            try:
                df_raw = pd.read_csv(path, dtype=str, engine="python", header=None)
                if df_raw.shape[1] == 1:
                    df_raw = pd.read_csv(path, dtype=str, engine="python", header=None, sep=";")
            except Exception:
                df_raw = pd.read_csv(path, dtype=str, engine="python", header=None, sep=";")
        else:
            df_raw = pd.read_excel(path, header=None, engine="openpyxl")

        hdr = SGP_HEADER_ROW if SGP_HEADER_ROW is not None else _detect_header_row(df_raw)
        hdr = max(0, min(hdr, len(df_raw)-1))
        cols = df_raw.iloc[hdr].astype(str).tolist()
        df_raw = df_raw.iloc[hdr+1:].reset_index(drop=True)
        df_raw.columns = cols
        dfn = _norm_cols(df_raw)

        info(f"[SGP] Cabeçalho aplicado (linha 1-based {hdr+1}): {list(dfn.columns)}")

        doc_col = None
        for c in ["cpf_cnpj","cpfcnpj","cpf","cnpj","documento","doc"]:
            if c in dfn.columns: doc_col = c; break
        if doc_col is None:
            doc_col = _pick_doc_col(dfn)
            info(f"[SGP] doc detectado por heurística: {doc_col}")
        if not doc_col:
            warn("[SGP] Não foi possível detectar coluna de documento.")
            return pd.DataFrame(columns=["Documento","SGP_Status","Nome_SGP"])

        docs = dfn[doc_col].astype(str).map(_digits)
        mask = docs.str.len().isin([11,14])

        name_col = None
        for c in ["nome_razao_social","nome","razao_social","nome_fantasia","cliente","pessoa"]:
            if c in dfn.columns: name_col = c; break
        if name_col is None:
            name_col = pick_name_column(dfn, exclude={doc_col})
            if name_col: info(f"[SGP] 'nome' por heurística: {name_col}")

        tmp = pd.DataFrame({"Documento": docs[mask]})
        tmp["__nome"] = dfn.loc[mask, name_col].astype(str).str.strip() if name_col else ""
        tmp["__len"] = tmp["__nome"].str.len()
        tmp = tmp.sort_values(["Documento","__len"], ascending=[True, False]).drop_duplicates("Documento", keep="first")

        sgp = tmp[["Documento"]].copy()
        sgp["SGP_Status"] = "NEGATIVADO"   # presença
        sgp["Nome_SGP"]   = tmp["__nome"].fillna("")
        info(f"[SGP] Presença marcada como NEGATIVADO para {len(sgp)} documentos.")
        return sgp[["Documento","SGP_Status","Nome_SGP"]]

    except Exception as e:
        warn(f"[SGP] Falha inesperada: {e}")
        return pd.DataFrame(columns=["Documento","SGP_Status","Nome_SGP"])

# ===== SOA =====
def ler_soa() -> pd.DataFrame:
    """
    SOA_Status (apenas dados do SOA):
      - Ativas -> "Negativado"
      - Baixadas -> "Baixado"
      - Pendentes -> "Pendente"
      - Erros -> "erro"
      - Determinacao -> "Determinação"
    Se houver Ativas e Baixadas, decide pela MAIOR DATA entre elas (SOA apenas).
    Caso não haja Ativas/Baixadas, prioriza: Pendente > erro > Determinação.
    Também captura nome se existir nas planilhas do SOA e devolve como Nome_SOA (para fallback).
    """
    frames = []
    cat_map = {
        "Ativas":       "ATIVAS",
        "Baixadas":     "BAIXADAS",
        "Determinacao": "DETERMINACAO",
        "Erros":        "ERROS",
        "Pendentes":    "PENDENTES",
    }

    for aba, path in SOA_CSVS.items():
        if not path.exists():
            continue
        try:
            df = pd.read_csv(path, dtype=str, engine="python")
            if df.shape[1] == 1:
                df = pd.read_csv(path, dtype=str, engine="python", sep=";")
        except Exception:
            df = pd.read_csv(path, dtype=str, engine="python", sep=";")

        dfn = _norm_cols(df)

        # Documento
        doc_col = None
        for c in ["documento","cpf_cnpj","cpfcnpj","cpf","cnpj","doc"]:
            if c in dfn.columns: doc_col = c; break
        if doc_col is None:
            best, best_r = None, -1
            for c in dfn.columns:
                r = dfn[c].astype(str).map(lambda v: len(_digits(v))>=9).mean()
                if r > best_r: best_r, best = r, c
            doc_col = best
            info(f"[SOA/{aba}] doc detectado: {doc_col} (ratio={best_r:.2f})")

        # Data
        date_map = _parse_dates_from_df(dfn)
        best_date_col = _pick_best_date(date_map)
        soa_date = date_map.get(best_date_col, pd.Series(pd.NaT, index=dfn.index))

        # Nome (se existir na planilha)
        name_col = None
        for c in ["nome_razao_social","nome","razao_social","nome_fantasia","cliente","pessoa"]:
            if c in dfn.columns: name_col = c; break
        if name_col is None:
            name_col = pick_name_column(dfn, exclude={doc_col})

        categoria = cat_map.get(aba, aba.upper())

        tmp = pd.DataFrame({
            "Documento": dfn[doc_col].astype(str).map(_digits),
            "Categoria": categoria,
            "SOA_Data":  soa_date,
            "Nome":      dfn[name_col].astype(str).str.strip() if name_col else ""
        })
        tmp = tmp[_valid_doc_mask(tmp["Documento"])].copy()
        frames.append(tmp)

    if not frames:
        warn("Sem CSVs do SOA.")
        return pd.DataFrame(columns=["Documento","SOA_Status","SOA_Data","SOA_Qtd","Nome_SOA"])

    soa_all = pd.concat(frames, ignore_index=True)

    def decide_status(g: pd.DataFrame) -> pd.Series:
        qtd_total = len(g)

        def max_dt(cat):
            s = g.loc[g["Categoria"] == cat, "SOA_Data"]
            return pd.to_datetime(s, errors="coerce").max()

        has_A = (g["Categoria"] == "ATIVAS").any()
        has_B = (g["Categoria"] == "BAIXADAS").any()
        has_P = (g["Categoria"] == "PENDENTES").any()
        has_E = (g["Categoria"] == "ERROS").any()
        has_D = (g["Categoria"] == "DETERMINACAO").any()

        dt_A = max_dt("ATIVAS")       if has_A else pd.NaT
        dt_B = max_dt("BAIXADAS")     if has_B else pd.NaT
        dt_P = max_dt("PENDENTES")    if has_P else pd.NaT
        dt_E = max_dt("ERROS")        if has_E else pd.NaT
        dt_D = max_dt("DETERMINACAO") if has_D else pd.NaT

        # Nome do SOA (melhor disponível no grupo)
        nome_soa = _coalesce_name(*g.get("Nome","").astype(str).tolist())

        # 1) Ativas x Baixadas pela maior data
        if has_A or has_B:
            if pd.notna(dt_A) and pd.notna(dt_B):
                if dt_A >= dt_B:
                    return pd.Series({"SOA_Qtd": qtd_total, "SOA_Status": "NEGATIVADO",   "SOA_Data": dt_A, "Nome_SOA": nome_soa})
                else:
                    return pd.Series({"SOA_Qtd": qtd_total, "SOA_Status": "BAIXADO",      "SOA_Data": dt_B, "Nome_SOA": nome_soa})
            if has_A:
                return pd.Series({"SOA_Qtd": qtd_total, "SOA_Status": "NEGATIVADO",       "SOA_Data": dt_A if pd.notna(dt_A) else pd.NaT, "Nome_SOA": nome_soa})
            if has_B:
                return pd.Series({"SOA_Qtd": qtd_total, "SOA_Status": "BAIXADO",          "SOA_Data": dt_B if pd.notna(dt_B) else pd.NaT, "Nome_SOA": nome_soa})

        # 2) Sem Ativas/Baixadas -> Pendente > erro > Determinação
        if has_P:
            return pd.Series({"SOA_Qtd": qtd_total, "SOA_Status": "PENDENTE",             "SOA_Data": dt_P if pd.notna(dt_P) else pd.NaT, "Nome_SOA": nome_soa})
        if has_E:
            return pd.Series({"SOA_Qtd": qtd_total, "SOA_Status": "ERRO",                 "SOA_Data": dt_E if pd.notna(dt_E) else pd.NaT, "Nome_SOA": nome_soa})
        if has_D:
            return pd.Series({"SOA_Qtd": qtd_total, "SOA_Status": "DETERMINAÇÃO",         "SOA_Data": dt_D if pd.notna(dt_D) else pd.NaT, "Nome_SOA": nome_soa})

        # 3) Sem sinal
        return pd.Series({"SOA_Qtd": qtd_total, "SOA_Status": "", "SOA_Data": pd.NaT, "Nome_SOA": nome_soa})

    agg = soa_all.groupby("Documento", as_index=False).apply(decide_status)
    return agg[["Documento","SOA_Status","SOA_Data","SOA_Qtd","Nome_SOA"]]

# ===== Consolidação (sem cruzar datas entre fontes) =====
def _consolidar_row_nocross(row):
    cnm = str(row.get("CNM_Status","")).upper().strip()
    soa = str(row.get("SOA_Status","")).upper().strip()
    sgp = str(row.get("SGP_Status","")).upper().strip()

    if cnm and soa and {cnm, soa} == {"NEGATIVADO","BAIXADO"}:
        return "ERRO"
    if cnm: return cnm
    if soa: return soa
    if sgp: return "NEGATIVADO"
    return "---"

def unificar_e_validar(df_cnm, df_soa, df_sgp) -> pd.DataFrame:
    base = df_cnm.merge(df_soa, on="Documento", how="outer")
    base = base.merge(df_sgp, on="Documento", how="outer")

    base["Documento"] = base["Documento"].astype(str).map(_digits)
    base = base[_valid_doc_mask(base["Documento"])].copy()

    for c in ["Nome","Nome_SGP","Nome_SOA","Tipo","CNM_Status","SOA_Status","SGP_Status","CNM_Data","SOA_Data","CNM_Qtd","SOA_Qtd"]:
        if c not in base.columns: base[c] = pd.NA

    # Nome: sempre tentar preencher com o melhor disponível (CNM > SGP > SOA, escolhendo o mais completo)
    def _pick_nome_row(row):
        return _coalesce_name(row.get("Nome",""), row.get("Nome_SGP",""), row.get("Nome_SOA",""))
    base["Nome"] = base.apply(_pick_nome_row, axis=1)

    # Normalizações
    base["Tipo"] = base.get("Tipo","").fillna("").astype(str).apply(_tipo_com_acento)
    base["CNM_Status"] = base.get("CNM_Status","").fillna("").astype(str)
    base["SOA_Status"] = base.get("SOA_Status","").fillna("").astype(str)
    base["SGP_Status"] = base.get("SGP_Status","").fillna("").astype(str)

    base["CNM_Data"] = pd.to_datetime(base.get("CNM_Data", pd.NaT), errors="coerce")
    base["SOA_Data"] = pd.to_datetime(base.get("SOA_Data", pd.NaT), errors="coerce")

    base["CNM_Qtd"] = pd.to_numeric(base.get("CNM_Qtd", 0), errors="coerce").fillna(0).astype(int)
    base["SOA_Qtd"] = pd.to_numeric(base.get("SOA_Qtd", 0), errors="coerce").fillna(0).astype(int)

    base["Consolidado"] = base.apply(_consolidar_row_nocross, axis=1)

    sem_sinal = (
        (base["CNM_Status"]=="") & (base["SOA_Status"]=="") & (base["SGP_Status"]=="") &
        (base["CNM_Qtd"]==0) & (base["SOA_Qtd"]==0)
    )
    base = base[~sem_sinal].copy()

    finais = ["Documento","Nome","SOA_Status","SGP_Status","CNM_Status","Tipo","Consolidado","CNM_Qtd","SOA_Qtd"]
    for c in finais:
        if c not in base.columns: base[c] = ""
    for c in ["Documento","Nome","SOA_Status","SGP_Status","CNM_Status","Tipo","Consolidado"]:
        base[c] = base[c].replace("", "---").fillna("---")

    base = base[finais].sort_values(["Documento","Nome"]).reset_index(drop=True)
    return base

# ===== Auditoria =====
def auditar_preenchimento(df_final: pd.DataFrame):
    info("Auditoria de preenchimento (vazios/'---' e valores inválidos):")
    alvo = ["Documento","Nome","SOA_Status","SGP_Status","CNM_Status","Tipo","Consolidado","CNM_Qtd","SOA_Qtd"]
    for c in alvo:
        vazios = (df_final[c].astype(str).isin(["", "---"])).sum()
        print(f" - {c:12s}: vazios/--- = {vazios}")

# ===== Main =====
def main():
    info(f"Download dir: {DOWNLOAD_DIR}")
    info(f"Output dir  : {OUTPUT_DIR}")

    df_cnm = ler_cnm()
    df_sgp = ler_sgp()
    df_soa = ler_soa()

    df_cnm = df_cnm if isinstance(df_cnm, pd.DataFrame) else pd.DataFrame(columns=["Documento","Nome","Tipo","CNM_Status","CNM_Qtd","CNM_Data"])
    df_sgp = df_sgp if isinstance(df_sgp, pd.DataFrame) else pd.DataFrame(columns=["Documento","SGP_Status","Nome_SGP"])
    df_soa = df_soa if isinstance(df_soa, pd.DataFrame) else pd.DataFrame(columns=["Documento","SOA_Status","SOA_Data","SOA_Qtd","Nome_SOA"])

    info(f"CNM linhas: {len(df_cnm)} | SGP: {len(df_sgp)} | SOA: {len(df_soa)}")

    df_final = unificar_e_validar(df_cnm, df_soa, df_sgp)

    auditar_preenchimento(df_final)

    with pd.ExcelWriter(OUT_XLSX, engine=ENGINE_WRITE) as writer:
        df_final.to_excel(writer, index=False, sheet_name="Dashboard")
        ws = writer.sheets["Dashboard"]
        widths = {
            "Documento":18, "Nome":36, "SOA_Status":14, "SGP_Status":14,
            "CNM_Status":14, "Tipo":12, "Consolidado":14,
            "CNM_Qtd":10, "SOA_Qtd":10
        }
        if ENGINE_WRITE == "xlsxwriter":
            for i, col in enumerate(df_final.columns):
                ws.set_column(i, i, widths.get(col, 16))
        else:
            from openpyxl.utils import get_column_letter
            for i, col in enumerate(df_final.columns, start=1):
                ws.column_dimensions[get_column_letter(i)].width = widths.get(col, 16)

    info(f"✅ Arquivo gerado: {OUT_XLSX}")

if __name__ == "__main__":
    main()
