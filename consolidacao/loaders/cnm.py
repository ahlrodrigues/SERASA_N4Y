import numpy as np
import pandas as pd
from pathlib import Path
from ..config import CNM_STD_XLSX, CNM_RAW_XLSX
from ..utils import (
    info, warn, _norm_cols, _digits, _valid_doc_mask,
    _parse_dates_from_df, _pick_best_date, _coalesce_name,
    _detect_header_row, _norm_text, pick_name_column
)

USAR_BAIXA = True  # EXCLUSAO -> "BAIXA" (se quiser "BAIXADO", use False)

OP_SUBSTR = ("INCLU", "EXCLU", "BAIX", "PROCESS")  # texto típico do campo operacional

def _tipo_com_acento(tipo_raw: str) -> str:
    t = (_norm_text(tipo_raw) or "").upper()
    if "INCLU" in t: return "INCLUSÃO"
    if "EXCLU" in t or "BAIX" in t: return "EXCLUSÃO"
    return "---"

def _status_por_tipo(tipo_raw: str) -> str:
    t = (_norm_text(tipo_raw) or "").upper()
    if "INCLU" in t: return "NEGATIVADO"
    if "EXCLU" in t or "BAIX" in t: return "BAIXA" if USAR_BAIXA else "BAIXADO"
    if "PROCESS" in t: return ""  # ignora PROCESSANDO
    return "ERRO" if t else ""

def _looks_operational(vals_norm: pd.Series) -> bool:
    if vals_norm.empty: return False
    sample = vals_norm.dropna().astype(str).head(200)
    if sample.empty: return False
    hit = sample.str.contains("|".join(OP_SUBSTR), regex=True).any()
    return bool(hit)

def _detect_tipo_column(dfn: pd.DataFrame) -> str | None:
    header_tokens = ("tipo", "status", "situacao", "situação", "operacao", "operação",
                     "evento", "acao", "ação", "mov", "movimento", "ocorrencia", "ocorrência", "tp")
    best, best_score = None, -1.0
    for c in dfn.columns:
        name = (_norm_text(c) or "").lower()
        hscore = sum(tok in name for tok in header_tokens) * 2.0
        vals_norm = dfn[c].astype(str).map(lambda x: (_norm_text(x) or "").upper())
        vscore = 0.0
        if not vals_norm.empty:
            frac = (
                vals_norm.str.contains("INCLU").mean()
              + vals_norm.str.contains("EXCLU").mean()
              + vals_norm.str.contains("BAIX").mean()
              + vals_norm.str.contains("PROCESS").mean()
            )
            vscore = float(frac) * 5.0
        score = hscore + vscore
        if score > best_score:
            best_score, best = score, c
    return best if best_score > 0 else None

def _last_nonempty(s: pd.Series, default=""):
    s2 = s.dropna().astype(str).replace("", np.nan).dropna()
    return s2.iloc[-1] if len(s2) else default

def ler_cnm() -> pd.DataFrame:
    # 1) Caminho standard (já normalizado)
    if CNM_STD_XLSX.exists():
        info(f"Lendo CNM standard: {CNM_STD_XLSX}")
        df = pd.read_excel(CNM_STD_XLSX)
        df["Documento"] = df["Documento"].astype(str).map(_digits)
        df = df[_valid_doc_mask(df["Documento"])].copy()
        df["Nome"] = df.get("Nome", "").astype(str).str.strip()

        # ✅ Troca Tipo -> CNM_TIPO
        df["CNM_TIPO"] = df.get("CNM_TIPO",
                          df.get("Tipo", "---")).apply(_tipo_com_acento)

        if "CNM_Status" not in df.columns or df["CNM_Status"].isna().all():
            def _map_tipo_to_status(x: str) -> str:
                if x == "INCLUSÃO":  return "NEGATIVADO"
                if x == "EXCLUSÃO":  return "BAIXA" if USAR_BAIXA else "BAIXADO"
                return "ERRO" if (x and x != "---") else ""
            df["CNM_Status"] = df["CNM_TIPO"].map(_map_tipo_to_status)

        agg = (
            df.assign(CNM_Data=pd.NaT)
              .groupby("Documento", as_index=False)
              .agg(
                  CNM_Qtd=("Documento", "size"),
                  Nome=("Nome", lambda s: _coalesce_name(*s)),
                  CNM_TIPO=("CNM_TIPO", _last_nonempty),
                  CNM_Status=("CNM_Status", _last_nonempty),
                  CNM_Data=("CNM_Data", "max")
              )
        )
        return agg[["Documento","Nome","CNM_TIPO","CNM_Status","CNM_Qtd","CNM_Data"]]

    # 2) Caminho bruto
    if not CNM_RAW_XLSX.exists():
        warn(f"CNM não encontrado: {CNM_STD_XLSX} nem {CNM_RAW_XLSX}")
        return pd.DataFrame(columns=["Documento","Nome","CNM_TIPO","CNM_Status","CNM_Qtd","CNM_Data"])

    info(f"Lendo CNM bruto: {CNM_RAW_XLSX}")
    df_raw = pd.read_excel(CNM_RAW_XLSX, header=None)
    header_row = _detect_header_row(df_raw)
    df_raw.columns = df_raw.iloc[header_row].astype(str).tolist()
    df_raw = df_raw.iloc[header_row+1:].reset_index(drop=True)
    dfn = _norm_cols(df_raw)

    aliases = {
        "documento": ["documento","cpf_cnpj","cpfcnpj","cpf","cnpj","doc","unnamed_3","unnamed_2"],
        "nome":      ["nome","nome_razao_social","razao_social","nome_fantasia","cliente","pessoa","nome_razao","razao"],
        "tipo":      ["tipo","tipo_","status","situacao","situação","operacao","operação","evento",
                      "acao","ação","movimento","mov","ocorrencia","ocorrência","tp","tipo_movimento"]
    }
    def pick(keys):
        for k in keys:
            if k in dfn.columns:
                return k
        return None

    col_doc  = pick(aliases["documento"])
    col_nome = pick(aliases["nome"])
    col_tipo = pick(aliases["tipo"])

    if col_doc is None:
        best, best_r = None, -1
        for c in dfn.columns:
            r = dfn[c].astype(str).map(lambda v: len(_digits(v)) in (11,14)).mean()
            if r > best_r:
                best_r, best = r, c
        col_doc = best
        info(f"[CNM] doc detectado por heurística: {col_doc} (ratio={best_r:.2f})")

    if col_nome is None:
        col_nome = pick_name_column(dfn, exclude={c for c in [col_doc, col_tipo] if c})
        if not col_nome:
            dfn["nome"] = ""
            col_nome = "nome"
            warn("[CNM] Nenhuma coluna de nome encontrada; criando 'nome' vazia.")
        else:
            info(f"[CNM] 'nome' por heurística: {col_nome}")

    # Detecta coluna operacional por conteúdo se necessário
    if col_tipo:
        vals_norm = dfn[col_tipo].astype(str).map(lambda x: (_norm_text(x) or "").upper())
        if not _looks_operational(vals_norm):
            col_tipo = None
    if not col_tipo:
        detected = _detect_tipo_column(dfn)
        if detected:
            info(f"[CNM] Coluna de tipo ajustada por conteúdo: {detected}")
            col_tipo = detected
        else:
            warn("[CNM] Nenhuma coluna operacional detectada; 'CNM_TIPO' ficará '---'.")
            dfn["tipo"] = ""
            col_tipo = "tipo"

    # Log amostra
    tipos_norm = dfn[col_tipo].astype(str).map(lambda x: (_norm_text(x) or "").upper())
    uniq_preview = sorted(tipos_norm.dropna().unique().tolist())[:20]
    info(f"[CNM] (coluna '{col_tipo}') amostra de TIPO normalizado: {uniq_preview}")

    # Datas
    date_map = _parse_dates_from_df(dfn)
    cnm_incl_col = next((c for c in date_map if "inclu" in c), None)
    cnm_excl_col = next((c for c in date_map if "exclu" in c or "baix" in c), None)
    best_date_col = _pick_best_date(date_map)

    # Base temporária (✅ usando CNM_TIPO)
    tmp = pd.DataFrame({
        "Documento": dfn[col_doc].astype(str).map(_digits),
        "Nome":      dfn[col_nome].astype(str).str.strip(),
        "CNM_TIPO":  tipos_norm.map(_tipo_com_acento),
        "CNM_Status": tipos_norm.map(_status_por_tipo),
    })

    # Datas por tipo — só INCLUSÃO/EXCLUSÃO; fallback para melhor coluna
    if cnm_incl_col or cnm_excl_col:
        cnm_date = pd.Series(pd.NaT, index=dfn.index)
        if cnm_incl_col is not None:
            cnm_date[tmp["CNM_TIPO"].eq("INCLUSÃO")] = pd.to_datetime(date_map[cnm_incl_col], errors="coerce")
        if cnm_excl_col is not None:
            cnm_date[tmp["CNM_TIPO"].eq("EXCLUSÃO")] = pd.to_datetime(date_map[cnm_excl_col], errors="coerce")
        mask_unk = tmp["CNM_TIPO"].eq("---")
        if mask_unk.any() and best_date_col is not None:
            cnm_date[mask_unk] = pd.to_datetime(date_map[best_date_col], errors="coerce")
    else:
        cnm_date = pd.to_datetime(date_map.get(best_date_col, pd.Series(pd.NaT, index=dfn.index)), errors="coerce")

    tmp["CNM_Data"] = cnm_date

    # Filtra documentos válidos
    tmp = tmp[_valid_doc_mask(tmp["Documento"])].copy()

    # Agrega por Documento (prioriza data mais recente)
    def agg_cnm(g: pd.DataFrame) -> pd.Series:
        g = g.sort_values("CNM_Data")
        qtd = len(g)
        nome = _coalesce_name(*g["Nome"])
        if g["CNM_Data"].notna().any():
            last = g.loc[g["CNM_Data"].idxmax()]
        else:
            last = g.iloc[-1].copy()
            st = g["CNM_Status"].replace("", np.nan).dropna()
            tp = g["CNM_TIPO"].replace("", np.nan).dropna()
            if len(st): last["CNM_Status"] = st.iloc[-1]
            if len(tp): last["CNM_TIPO"] = tp.iloc[-1]
            last["CNM_Data"] = pd.NaT
        return pd.Series({
            "CNM_Qtd": qtd,
            "Nome": nome,
            "CNM_Status": str(last["CNM_Status"]),
            "CNM_TIPO": str(last["CNM_TIPO"]),
            "CNM_Data": last["CNM_Data"]
        })

    agg = tmp.groupby("Documento").apply(agg_cnm).reset_index()
    return agg[["Documento","Nome","CNM_TIPO","CNM_Status","CNM_Qtd","CNM_Data"]]
