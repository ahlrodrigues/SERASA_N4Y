import pandas as pd
from pathlib import Path
from ..config import SGP_XLSX, SGP_HEADER_ROW, DOWNLOAD_DIR
from ..utils import info, warn, _norm_cols, _digits, _valid_doc_mask, pick_name_column, find_latest

def _pick_doc_col(df: pd.DataFrame) -> str:
    best, best_score = None, -1
    for c in df.columns:
        s = df[c].astype(str).map(_digits)
        ratio = ((s.str.len()==11) | (s.str.len()==14)).mean()
        uniq  = s.nunique(dropna=True)/max(1, (s!="").sum())
        score = ratio*0.9 + uniq*0.1
        if score > best_score:
            best_score, best = score, c
    return best

def ler_sgp() -> pd.DataFrame:
    """
    Retorna apenas os documentos presentes no SGP com status 'NEGATIVADO'.
    A marcação '---' para ausentes é feita na consolidação com merge + fillna.
    Saída: colunas ['Documento','SGP_status','Nome_SGP']
    """
    try:
        path = SGP_XLSX if SGP_XLSX.exists() else find_latest(
            ["Relatorio_SGP.*", "*SGP*.xlsx", "*sgp*.xlsx", "*SGP*.xls", "*sgp*.xls", "*SGP*.csv", "*sgp*.csv"]
        )
        if not path or not Path(path).exists():
            warn(f"SGP não encontrado: {SGP_XLSX} (e nenhum *SGP*.xlsx/.xls/.csv em {DOWNLOAD_DIR})")
            return pd.DataFrame(columns=["Documento","SGP_status","Nome_SGP"])

        info(f"Lendo SGP de: {path}")

        # leitura robusta
        ext = str(path).lower()
        if ext.endswith(".csv"):
            try:
                df_raw = pd.read_csv(path, dtype=str, engine="python", header=None)
                if df_raw.shape[1] == 1:
                    df_raw = pd.read_csv(path, dtype=str, engine="python", header=None, sep=";")
            except Exception:
                try:
                    df_raw = pd.read_csv(path, dtype=str, engine="python", header=None, sep=";", encoding="latin-1")
                except Exception:
                    df_raw = pd.read_csv(path, dtype=str, engine="python", header=None, encoding_errors="ignore")
        else:
            # Deixe o pandas escolher o engine (suporta .xlsx e .xls se o engine estiver instalado)
            df_raw = pd.read_excel(path, header=None)

        # autodetect do cabeçalho
        if SGP_HEADER_ROW is not None:
            hdr = SGP_HEADER_ROW
        else:
            def _detect_header_row_sgp(df):
                tokens = {
                    "documento","cpf","cnpj","cpf_cnpj","cpfcnpj",
                    "nome","razao","razao_social","fantasia","cliente","pessoa",
                    "status","situacao","situação","operacao","tipo","data","dt"
                }
                best_i, best_score = 0, -1
                max_scan = min(120, len(df))
                for i in range(max_scan):
                    row = df.iloc[i].astype(str).fillna("").tolist()
                    rown = [str(x).lower() for x in row]
                    nonempty = sum(1 for x in rown if x and x != "nan")
                    hits = sum(1 for x in rown if any(t in x for t in tokens))
                    score = hits * 4 + nonempty
                    if score > best_score:
                        best_score, best_i = score, i
                return best_i
            hdr = _detect_header_row_sgp(df_raw)

        hdr = max(0, min(hdr, len(df_raw)-1))
        cols = df_raw.iloc[hdr].astype(str).tolist()
        df_raw = df_raw.iloc[hdr+1:].reset_index(drop=True)
        df_raw.columns = cols
        dfn = _norm_cols(df_raw)

        info(f"[SGP] Cabeçalho (1-based {hdr+1}): {list(dfn.columns)}")

        # coluna documento
        doc_col = None
        for c in ["cpf_cnpj","cpfcnpj","cpf","cnpj","documento","doc"]:
            if c in dfn.columns:
                doc_col = c; break
        if doc_col is None:
            doc_col = _pick_doc_col(dfn)
            info(f"[SGP] doc detectado por heurística: {doc_col}")

        if not doc_col:
            warn("[SGP] Não foi possível detectar coluna de documento.")
            return pd.DataFrame(columns=["Documento","SGP_status","Nome_SGP"])

        # normaliza docs
        docs = dfn[doc_col].astype(str).map(_digits)
        mask_valid = _valid_doc_mask(docs)  # usa sua função padrão do projeto

        # coluna nome
        name_col = None
        for c in ["nome_razao_social","nome","razao_social","nome_fantasia","cliente","pessoa"]:
            if c in dfn.columns:
                name_col = c; break
        if name_col is None:
            name_col = pick_name_column(dfn, exclude={doc_col})
            if name_col:
                info(f"[SGP] 'nome' por heurística: {name_col}")

        tmp = pd.DataFrame({
            "Documento": docs[mask_valid].astype(str)
        })
        if name_col:
            tmp["__nome"] = dfn.loc[mask_valid, name_col].astype(str).str.strip()
        else:
            tmp["__nome"] = ""

        # Para cada Documento, prioriza o nome mais completo (maior comprimento)
        tmp["__len"] = tmp["__nome"].str.len()
        tmp = (
            tmp.sort_values(["Documento","__len"], ascending=[True, False])
               .drop_duplicates("Documento", keep="first")
        )

        sgp = tmp[["Documento"]].copy()
        sgp["SGP_status"] = "NEGATIVADO"  # presença no SGP
        sgp["Nome_SGP"]   = tmp["__nome"].fillna("").astype(str)

        # garante tipos/limpeza
        sgp["Documento"]  = sgp["Documento"].astype(str)
        sgp["SGP_status"] = sgp["SGP_status"].astype(str)

        info(f"[SGP] Linhas após limpeza: {len(sgp)}")
        return sgp[["Documento","SGP_status","Nome_SGP"]]

    except Exception as e:
        warn(f"[SGP] Falha inesperada: {e}")
        return pd.DataFrame(columns=["Documento","SGP_status","Nome_SGP"])
