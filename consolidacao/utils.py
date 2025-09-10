import os, re, glob, unicodedata
import numpy as np
import pandas as pd
from pathlib import Path

from .config import DOWNLOAD_DIR

def info(m): print(f"[INFO] {m}")
def warn(m): print(f"[WARN] {m}")

def _norm_text(s):
    if pd.isna(s): return ""
    s = str(s).strip()
    s = unicodedata.normalize("NFKD", s).encode("ascii", "ignore").decode("ascii")
    return s

def _norm_cols(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df.columns = [
        re.sub(r"[^a-z0-9]+", "_", _norm_text(c).lower()).strip("_")
        for c in df.columns
    ]
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

def pick_name_column(df: pd.DataFrame, exclude=()):
    import re as _re
    exclude = set(exclude or ())
    candidates = []
    for c in df.columns:
        if c in exclude: continue
        cl = str(c).lower()
        if any(tok in cl for tok in [
            "usuario","user","login","email","e_mail","mail",
            "operacao","tipo","status","situacao","situação",
            "data","hora","dt","cod","codigo","código","id",
            "documento","cpf","cnpj","doc","chave","hash","origem"
        ]):
            continue
        s = df[c].astype(str).fillna("")
        if (s == "").mean() > 0.90: continue
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
