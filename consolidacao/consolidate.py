# === consolidado.py (trechos principais) =====================================
import pandas as pd
from pathlib import Path
from .config import DOWNLOAD_DIR, OUTPUT_DIR, OUT_XLSX, ENGINE_WRITE, CONSOLIDADO_STYLE
from .utils import info, _digits, _valid_doc_mask, _coalesce_name
from .loaders.cnm import ler_cnm, _tipo_com_acento
from .loaders.sgp import ler_sgp
from .loaders.soa import ler_soa

def _format_status_for_output(status: str) -> str:
    if status in (None, "", "---"):
        return status or ""
    if CONSOLIDADO_STYLE.upper() == "SOA":
        key = str(status).upper()
        mapping = {
            "NEGATIVADO": "Negativado",
            "BAIXADO": "Baixado",
            "BAIXA": "Baixa",
            "PENDENTE": "Pendente",
            "ERRO": "erro",
            "DETERMINACAO": "Determinação",
            "DETERMINAÇÃO": "Determinação",
        }
        return mapping.get(key, status)
    return str(status).upper()

def _consolidar_row_nocross(row) -> str:
    cnm = str(row.get("CNM_Status", "")).strip().upper()
    soa = str(row.get("SOA_Status", "")).strip().upper()
    sgp = str(row.get("SGP_Status", "")).strip().upper()

    if soa == "BAIXADO" and cnm == "NEGATIVADO" and sgp == "NEGATIVADO":
        return _format_status_for_output("NEGATIVADO")

    if cnm and soa and {cnm, soa} == {"NEGATIVADO", "BAIXADO"} or {cnm, soa} == {"NEGATIVADO", "BAIXA"}:
        return _format_status_for_output("ERRO")

    if cnm:
        return _format_status_for_output(cnm)
    if soa:
        return _format_status_for_output(soa)
    if sgp == "NEGATIVADO":
        return _format_status_for_output("ERRO")
    return "---"

def unificar_e_validar(df_cnm: pd.DataFrame, df_soa: pd.DataFrame, df_sgp: pd.DataFrame) -> pd.DataFrame:
    base = df_cnm.merge(df_soa, on="Documento", how="outer")
    base = base.merge(df_sgp, on="Documento", how="outer")

    # 🔧 normaliza nomes vindos de loaders (maiúsc/minúsc e compatibilidade antiga)
    rename_map = {}
    for col in list(base.columns):
        low = col.lower()
        if low == "sgp_status" and "SGP_Status" not in base.columns:
            rename_map[col] = "SGP_Status"
        if low == "soa_status" and "SOA_Status" not in base.columns:
            rename_map[col] = "SOA_Status"
        if low == "cnm_status" and "CNM_Status" not in base.columns:
            rename_map[col] = "CNM_Status"
        # ✅ aceita "Tipo" legado como CNM_TIPO
        if low == "tipo" and "CNM_TIPO" not in base.columns:
            rename_map[col] = "CNM_TIPO"
        if low == "cnm_tipo" and "CNM_TIPO" not in base.columns:
            rename_map[col] = "CNM_TIPO"
        if low == "nome_sgp" and "Nome_SGP" not in base.columns:
            rename_map[col] = "Nome_SGP"
        if low == "nome_soa" and "Nome_SOA" not in base.columns:
            rename_map[col] = "Nome_SOA"
    if rename_map:
        base = base.rename(columns=rename_map)

    base["Documento"] = base["Documento"].astype(str).map(_digits)
    base = base[_valid_doc_mask(base["Documento"])].copy()

    for c in ["Nome","Nome_SGP","Nome_SOA","CNM_TIPO","CNM_Status","SOA_Status","SGP_Status",
              "CNM_Data","SOA_Data","CNM_Qtd","SOA_Qtd"]:
        if c not in base.columns:
            base[c] = pd.NA

    base["Nome"] = base.apply(
        lambda r: _coalesce_name(r.get("Nome",""),
                                 r.get("Nome_SGP",""),
                                 r.get("Nome_SOA","")),
        axis=1
    )

    # ✅ normaliza CNM_TIPO (mantém compatível se vier vazio)
    base["CNM_TIPO"]   = base.get("CNM_TIPO","").fillna("").astype(str).apply(_tipo_com_acento)
    base["CNM_Status"] = base.get("CNM_Status","").fillna("").astype(str)
    base["SOA_Status"] = base.get("SOA_Status","").fillna("").astype(str)
    base["SGP_Status"] = base.get("SGP_Status","").fillna("").astype(str)

    base["CNM_Data"] = pd.to_datetime(base.get("CNM_Data", pd.NaT), errors="coerce")
    base["SOA_Data"] = pd.to_datetime(base.get("SOA_Data", pd.NaT), errors="coerce")

    base["CNM_Qtd"] = pd.to_numeric(base.get("CNM_Qtd", 0), errors="coerce").fillna(0).astype(int)
    base["SOA_Qtd"] = pd.to_numeric(base.get("SOA_Qtd", 0), errors="coerce").fillna(0).astype(int)

    base["Consolidado"] = base.apply(_consolidar_row_nocross, axis=1)

    for col in ["SOA_Status", "CNM_Status", "SGP_Status", "Consolidado"]:
        base[col] = base[col].map(_format_status_for_output)

    sem_sinal = (
        (base["CNM_Status"].fillna("")=="") &
        (base["SOA_Status"].fillna("")=="") &
        (base["SGP_Status"].fillna("")=="") &
        (base["CNM_Qtd"]==0) & (base["SOA_Qtd"]==0)
    )
    base = base[~sem_sinal].copy()

    # ✅ colunas finais agora usam CNM_TIPO
    finais = ["Documento","Nome","SOA_Status","SGP_Status","CNM_Status","CNM_TIPO","Consolidado","CNM_Qtd","SOA_Qtd"]
    for c in finais:
        if c not in base.columns:
            base[c] = ""

    for c in ["Documento","Nome","SOA_Status","SGP_Status","CNM_Status","CNM_TIPO","Consolidado"]:
        base[c] = base[c].replace("", "---").fillna("---")

    base = base[finais].sort_values(["Documento","Nome"]).reset_index(drop=True)
    return base

def _salvar_excel(df_final: pd.DataFrame) -> None:
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    with pd.ExcelWriter(OUT_XLSX, engine=ENGINE_WRITE) as writer:
        df_final.to_excel(writer, index=False, sheet_name="Dashboard")
        ws = writer.sheets["Dashboard"]
        widths = {
            "Documento": 18, "Nome": 36,
            "SOA_Status": 14, "SGP_Status": 14, "CNM_Status": 14,
            "CNM_TIPO": 12, "Consolidado": 14,
            "CNM_Qtd": 10, "SOA_Qtd": 10
        }
        if ENGINE_WRITE == "xlsxwriter":
            for i, col in enumerate(df_final.columns):
                ws.set_column(i, i, widths.get(col, 16))
        else:
            from openpyxl.utils import get_column_letter
            for i, col in enumerate(df_final.columns, start=1):
                ws.column_dimensions[get_column_letter(i)].width = widths.get(col, 16)


# ----------------------------------- main ------------------------------------
def main():
    DOWNLOAD_DIR.mkdir(parents=True, exist_ok=True)
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

    df_cnm = ler_cnm()
    df_sgp = ler_sgp()
    df_soa = ler_soa()

    info(f"CNM linhas: {len(df_cnm)} | SGP: {len(df_sgp)} | SOA: {len(df_soa)}")
    if len(df_sgp) == 0:
        info("[SGP] AVISO: 0 linhas detectadas. Verifique cabeçalho/colunas do arquivo SGP.")

    df_final = unificar_e_validar(df_cnm, df_soa, df_sgp)
    _salvar_excel(df_final)
    info(f"✅ Arquivo gerado: {OUT_XLSX}")


if __name__ == "__main__":
    main()