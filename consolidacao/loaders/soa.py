import pandas as pd
import numpy as np
from ..config import SOA_CSVS
from ..utils import info, warn, _norm_cols, _digits, _valid_doc_mask, _parse_dates_from_df, _pick_best_date, _coalesce_name

def ler_soa() -> pd.DataFrame:
    """
    SOA_Status:
      - Ativas -> "Negativado"
      - Baixadas -> "Baixado"
      - Pendentes -> "Pendente"
      - Erros -> "erro"
      - Determinacao -> "Determinação"
    Ativas x Baixadas: decide pela maior data (SOA apenas).
    Sem Ativas/Baixadas: Pendente > erro > Determinação.
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
            try:
                df = pd.read_csv(path, dtype=str, engine="python", sep=";", encoding="latin-1")
            except Exception:
                df = pd.read_csv(path, dtype=str, engine="python", encoding_errors="ignore")

        dfn = _norm_cols(df)

        # Documento
        doc_col = None
        for c in ["documento","cpf_cnpj","cpfcnpj","cpf","cnpj","doc"]:
            if c in dfn.columns: doc_col = c; break
        if doc_col is None:
            best, best_r = None, -1
            for c in dfn.columns:
                r = dfn[c].astype(str).map(lambda v: len(_digits(v)) >= 9).mean()
                if r > best_r: best_r, best = r, c
            doc_col = best
            info(f"[SOA/{aba}] doc detectado: {doc_col} (ratio={best_r:.2f})")

        # Data
        date_map = _parse_dates_from_df(dfn)
        best_date_col = _pick_best_date(date_map)
        soa_date = date_map.get(best_date_col, pd.Series(pd.NaT, index=dfn.index))

        # Nome (se existir)
        name_col = None
        for c in ["nome_razao_social","nome","razao_social","nome_fantasia","cliente","pessoa"]:
            if c in dfn.columns: name_col = c; break
        if name_col is None:
            from ..utils import pick_name_column
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

        nome_soa = _coalesce_name(*g.get("Nome", "").astype(str).tolist())

        if has_A or has_B:
            if pd.notna(dt_A) and pd.notna(dt_B):
                if dt_A >= dt_B:
                    return pd.Series({"SOA_Qtd": qtd_total, "SOA_Status": "Negativado", "SOA_Data": dt_A, "Nome_SOA": nome_soa})
                else:
                    return pd.Series({"SOA_Qtd": qtd_total, "SOA_Status": "Baixado", "SOA_Data": dt_B, "Nome_SOA": nome_soa})
            if has_A:
                return pd.Series({"SOA_Qtd": qtd_total, "SOA_Status": "Negativado", "SOA_Data": dt_A if pd.notna(dt_A) else pd.NaT, "Nome_SOA": nome_soa})
            if has_B:
                return pd.Series({"SOA_Qtd": qtd_total, "SOA_Status": "Baixado", "SOA_Data": dt_B if pd.notna(dt_B) else pd.NaT, "Nome_SOA": nome_soa})

        if has_P:
            return pd.Series({"SOA_Qtd": qtd_total, "SOA_Status": "Pendente", "SOA_Data": dt_P if pd.notna(dt_P) else pd.NaT, "Nome_SOA": nome_soa})
        if has_E:
            return pd.Series({"SOA_Qtd": qtd_total, "SOA_Status": "erro", "SOA_Data": dt_E if pd.notna(dt_E) else pd.NaT, "Nome_SOA": nome_soa})
        if has_D:
            return pd.Series({"SOA_Qtd": qtd_total, "SOA_Status": "Determinação", "SOA_Data": dt_D if pd.notna(dt_D) else pd.NaT, "Nome_SOA": nome_soa})

        return pd.Series({"SOA_Qtd": qtd_total, "SOA_Status": "", "SOA_Data": pd.NaT, "Nome_SOA": nome_soa})

    agg = soa_all.groupby("Documento", as_index=False).apply(decide_status)
    return agg[["Documento","SOA_Status","SOA_Data","SOA_Qtd","Nome_SOA"]]
