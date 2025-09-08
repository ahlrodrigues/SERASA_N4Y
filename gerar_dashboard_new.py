# -*- coding: utf-8 -*-
"""
GERADOR DE DASHBOARD UNIFICADO (CNM x SOA x SGP)
- Preserva as colunas "Operação" e "tipo" do CNM
- Aplica regra consolidada pedida:
  Se SGP=NEGATIVADO e CNM='---' e SOA='---' => Consolidado='ERRO'
- Remove linhas PROCESSANDO do CNM e PENDENTES do SOA
- Gera HTML com filtro por Status e subtítulo dinâmico
"""

import os
import re
import sys
import glob
import shutil
import unicodedata
from pathlib import Path

import pandas as pd


# =========================
# ======== CONFIG =========
# =========================
DOWNLOAD_DIR = "./download"
OUTPUT_DIR   = "./output"

# Arquivos padrão (ajuste conforme seu ambiente)
CNM_FILE = os.path.join(DOWNLOAD_DIR, "Relatorio_CNM.xlsx")
SGP_FILE = os.path.join(DOWNLOAD_DIR, "Relatorio_SGP.xlsx")

# SOA: ler múltiplos CSVs se existirem
SOA_CSVS = {
    "Ativas":       os.path.join(DOWNLOAD_DIR, "Ativas.csv"),
    "Baixadas":     os.path.join(DOWNLOAD_DIR, "Baixadas.csv"),
    "Determinacao": os.path.join(DOWNLOAD_DIR, "Determinacao.csv"),
    "Erros":        os.path.join(DOWNLOAD_DIR, "Erros.csv"),
    "Pendentes":    os.path.join(DOWNLOAD_DIR, "Pendentes.csv"),
}


# =========================
# ======= UTILS ===========
# =========================
def info(msg): print(f"[INFO] {msg}")
def warn(msg): print(f"[WARN] {msg}")
def err(msg):  print(f"[ERRO] {msg}", file=sys.stderr)

def _norm_text(s):
    if pd.isna(s):
        return ""
    s = str(s).strip()
    s = unicodedata.normalize("NFKD", s).encode("ascii", "ignore").decode("ascii")
    return s

def _norm_cols(df: pd.DataFrame) -> pd.DataFrame:
    new_cols = []
    for c in df.columns:
        base = _norm_text(c).lower().replace(" ", "_")
        new_cols.append(base)
    df.columns = new_cols
    return df

def _as_digits(s):
    return re.sub(r"\D", "", str(s) if s is not None else "")

def ensure_output_clean():
    Path(OUTPUT_DIR).mkdir(parents=True, exist_ok=True)
    # apaga tudo dentro de output/
    for p in Path(OUTPUT_DIR).glob("*"):
        if p.is_file():
            p.unlink()
        elif p.is_dir():
            shutil.rmtree(p)


# =========================
# ======== CNM =============
# =========================
def _detect_header_row(df_raw: pd.DataFrame, max_scan=15):
    """
    Tenta detectar a linha do cabeçalho: procura linhas onde
    apareçam termos usuais de documento/nome.
    """
    patterns = ["documento", "cpf", "cnpj", "cpf_cnpj", "nome", "razao", "fantasia"]
    for i in range(min(max_scan, len(df_raw))):
        row = df_raw.iloc[i].astype(str).str.lower().fillna("")
        joined = " ".join(row.tolist())
        if any(p in joined for p in patterns):
            return i
    return 0  # fallback


def ler_cnm(path_cnm: str) -> pd.DataFrame:
    if not os.path.exists(path_cnm):
        # tenta CSV genérico
        csv_guess = glob.glob(os.path.join(DOWNLOAD_DIR, "Relatorio_CNM.*"))
        if csv_guess:
            path_cnm = csv_guess[0]
        else:
            raise FileNotFoundError(f"Arquivo CNM não encontrado em {path_cnm}")

    info(f"Lendo CNM: {path_cnm}")
    if path_cnm.lower().endswith((".xls", ".xlsx")):
        df_raw = pd.read_excel(path_cnm, header=None)
        header_row = _detect_header_row(df_raw, max_scan=20)
        df_raw.columns = df_raw.iloc[header_row].astype(str).tolist()
        df_raw = df_raw.iloc[header_row+1:].reset_index(drop=True)
    else:
        # CSV com ; é comum, mas tentamos inteligentemente
        try:
            df_raw = pd.read_csv(path_cnm, sep=";", dtype=str, engine="python")
        except Exception:
            df_raw = pd.read_csv(path_cnm, sep=",", dtype=str, engine="python")

    df_raw = df_raw.copy()
    df = _norm_cols(df_raw)

    # aliases de colunas em CNM
    aliases = {
        "documento": ["documento","cpf_cnpj","cpfcnpj","cpf","cnpj","doc","unnamed:_3","unnamed:_2"],
        "nome": ["nome","nome_razao_social","razao_social","nome_fantasia","cliente","pessoa"],
        "operacao": ["operacao","operacao_","tipo_operacao"],
        "tipo": ["tipo","tipo_","status","situacao"]
    }

    def pick(keys):
        for k in keys:
            if k in df.columns:
                return k
        return None

    col_doc  = pick(aliases["documento"])
    col_nome = pick(aliases["nome"])
    col_op   = pick(aliases["operacao"])
    col_tipo = pick(aliases["tipo"])

    if col_doc is None:
        # Heurística: escolha a coluna com maior proporção de células numéricas longas
        best_col, best_ratio = None, -1
        for c in df.columns:
            series = df[c].astype(str).fillna("")
            digits = series.map(_as_digits)
            long_num = digits.map(lambda x: len(x) >= 9)  # CPF/CNPJ/Doc
            ratio = long_num.mean()
            if ratio > best_ratio:
                best_ratio = ratio
                best_col = c
        col_doc = best_col
        info(f"Coluna de documento detectada por amostragem: '{col_doc}' (ratio={best_ratio:.2f})")

    # garante todas
    if col_nome is None: df["nome"] = ""
    if col_op   is None: df["operacao"] = ""
    if col_tipo is None: df["tipo"] = ""

    col_nome = col_nome or "nome"
    col_op   = col_op or "operacao"
    col_tipo = col_tipo or "tipo"

    # normalizações
    doc_norm  = df[col_doc].astype(str).map(_as_digits).str.zfill(11)
    nome_norm = df[col_nome].astype(str).str.strip()

    # preserva "Operação" e "tipo" exatamente
    operacao_preserva = df[col_op].astype(str).str.strip() if col_op in df.columns else ""
    tipo_preserva     = df[col_tipo].astype(str).str.strip() if col_tipo in df.columns else ""

    out = pd.DataFrame({
        "Documento": doc_norm,
        "Nome": nome_norm,
        "Operação": operacao_preserva,
        "tipo": tipo_preserva,
    })

    # status do CNM baseado em "tipo"
    out["CNM_Status"] = out["tipo"].map(status_por_tipo_cnm)
    return out


def status_por_tipo_cnm(tipo: str) -> str:
    t = _norm_text(tipo).upper()
    if t == "INCLUSAO":
        return "NEGATIVADO"
    if t == "EXCLUSAO":
        return "BAIXADO"
    if t == "PROCESSANDO":
        return ""  # ignorar mais à frente
    return "ERRO" if t else ""


# =========================
# ======== SGP =============
# =========================
def ler_sgp(path_sgp: str) -> pd.DataFrame:
    if not os.path.exists(path_sgp):
        # tenta variantes
        guess = glob.glob(os.path.join(DOWNLOAD_DIR, "Relatorio_SGP.*"))
        if guess:
            path_sgp = guess[0]
        else:
            warn(f"Arquivo SGP não encontrado em {path_sgp}. Devolvendo vazio.")
            return pd.DataFrame(columns=["Documento","SGP_Status"])

    info(f"Lendo SGP: {path_sgp}")
    if path_sgp.lower().endswith((".xls", ".xlsx")):
        df = pd.read_excel(path_sgp)
    else:
        try:
            df = pd.read_csv(path_sgp, sep=";", dtype=str, engine="python")
        except Exception:
            df = pd.read_csv(path_sgp, sep=",", dtype=str, engine="python")

    df = _norm_cols(df)

    # Detectar doc e status
    # Documento costuma estar em 'cpf_cnpj' ou similar
    doc_col = None
    for c in ["cpf_cnpj","cpfcnpj","cpf","cnpj","documento","doc"]:
        if c in df.columns:
            doc_col = c
            break
    if doc_col is None:
        # heurística
        best_col, best_ratio = None, -1
        for c in df.columns:
            series = df[c].astype(str).fillna("")
            digits = series.map(_as_digits)
            long_num = digits.map(lambda x: len(x) >= 9)
            ratio = long_num.mean()
            if ratio > best_ratio:
                best_ratio = ratio
                best_col = c
        doc_col = best_col
        info(f"[SGP] Coluna de documento detectada por amostragem: '{doc_col}' (ratio={best_ratio:.2f})")

    # Status no SGP:
    # Você pode já ter um campo de status; se não, mapeie a partir de alguma coluna de "Tipo"
    status_col = None
    for c in ["status","situacao","tipo","resultado"]:
        if c in df.columns:
            status_col = c
            break

    out = pd.DataFrame()
    out["Documento"] = df[doc_col].astype(str).map(_as_digits)
    if status_col:
        s = df[status_col].astype(str).str.upper().str.strip()
        # normalizar para NEGATIVADO/BAIXADO/ERRO quando possível
        out["SGP_Status"] = s.replace({
            "INCLUSAO": "NEGATIVADO",
            "EXCLUSAO": "BAIXADO",
            "NEGATIVADO": "NEGATIVADO",
            "BAIXADO": "BAIXADO",
        })
        out.loc[~out["SGP_Status"].isin(["NEGATIVADO","BAIXADO"]), "SGP_Status"] = "ERRO"
    else:
        # Se não houver, assuma vazio (será '---' depois)
        out["SGP_Status"] = ""

    # consolidar por documento (último/primeiro válido)
    out = (out
           .groupby("Documento", as_index=False)["SGP_Status"]
           .agg(lambda x: next((v for v in x if v), "")))
    return out


# =========================
# ======== SOA =============
# =========================
def ler_soa(csv_map: dict) -> pd.DataFrame:
    frames = []
    for aba, path in csv_map.items():
        if not os.path.exists(path):
            continue
        info(f"Lendo arquivo SOA: {aba} -> {path}")
        try:
            df = pd.read_csv(path, dtype=str, engine="python")
        except Exception:
            # tenta ; como separador secundário
            df = pd.read_csv(path, dtype=str, engine="python", sep=";")
        df = _norm_cols(df)

        # detectar documento
        doc_col = None
        for c in ["documento","cpf_cnpj","cpfcnpj","cpf","cnpj","doc"]:
            if c in df.columns:
                doc_col = c
                break
        if doc_col is None:
            # heurística
            best_col, best_ratio = None, -1
            for c in df.columns:
                series = df[c].astype(str).fillna("")
                digits = series.map(_as_digits)
                long_num = digits.map(lambda x: len(x) >= 9)
                ratio = long_num.mean()
                if ratio > best_ratio:
                    best_ratio = ratio
                    best_col = c
            doc_col = best_col
            info(f"[SOA/{aba}] Coluna doc detectada: '{doc_col}' (ratio={best_ratio:.2f})")

        tmp = pd.DataFrame()
        tmp["Documento"] = df[doc_col].astype(str).map(_as_digits)

        if aba.lower() == "ativas":
            tmp["SOA_Status"] = "NEGATIVADO"
            tmp["SOA_Origem"] = "Ativas"
        elif aba.lower() == "baixadas":
            tmp["SOA_Status"] = "BAIXADO"
            tmp["SOA_Origem"] = "Baixadas"
        elif aba.lower() in ("determinacao","determinação"):
            tmp["SOA_Status"] = "ERRO"
            tmp["SOA_Origem"] = "Determinacao"
        elif aba.lower() == "erros":
            tmp["SOA_Status"] = "ERRO"
            tmp["SOA_Origem"] = "Erros"
        elif aba.lower() == "pendentes":
            # ignorar no final
            tmp["SOA_Status"] = ""
            tmp["SOA_Origem"] = "Pendentes"
        else:
            tmp["SOA_Status"] = ""
            tmp["SOA_Origem"] = aba

        frames.append(tmp)

    if not frames:
        warn("Nenhum CSV do SOA encontrado. Devolvendo vazio.")
        return pd.DataFrame(columns=["Documento","SOA_Status","SOA_Origem"])

    df_soa = pd.concat(frames, ignore_index=True)
    # consolida por documento: prioriza NEGATIVADO/BAIXADO, senão ERRO, senão vazio
    def pick_status(g):
        vals = list(g.dropna().astype(str))
        for pref in ["NEGATIVADO","BAIXADO","ERRO"]:
            if pref in vals:
                return pref
        return ""

    out = df_soa.groupby("Documento", as_index=False).agg({
        "SOA_Status": pick_status,
        "SOA_Origem": lambda x: next((v for v in x if v), "---")
    })
    return out


# =========================
# ===== CONSOLIDAÇÃO ======
# =========================
def consolidar_registro(row):
    """
    Espera no DF:
      - CNM_Status, SOA_Status, SGP_Status
      - CNM_Origem, SOA_Origem (para a regra especial)
    """
    cnm_status = str(row.get("CNM_Status", "")).strip().upper()
    soa_status = str(row.get("SOA_Status", "")).strip().upper()
    sgp_status = str(row.get("SGP_Status", "")).strip().upper()

    cnm_origem = str(row.get("CNM_Origem", "")).strip() or "---"
    soa_origem = str(row.get("SOA_Origem", "")).strip() or "---"

    # Regra especial solicitada:
    if sgp_status == "NEGATIVADO" and cnm_origem == "---" and soa_origem == "---":
        return "ERRO"

    statuses = {cnm_status, soa_status, sgp_status} - {""}
    if not statuses:
        return "---"

    if "ERRO" in statuses:
        return "ERRO"

    if {"NEGATIVADO","BAIXADO"}.issubset(statuses):
        return "ERRO"

    if len(statuses) == 1:
        return list(statuses)[0]

    # divergência sem regra explícita
    return "ERRO"


def construir_unificado(df_cnm, df_soa, df_sgp):
    # merges
    base = df_cnm.merge(df_soa, on="Documento", how="outer")
    base = base.merge(df_sgp, on="Documento", how="outer")

    # origem CNM: usa "Operação" se existir, senão '---'
    base["CNM_Origem"] = base["Operação"].where(base["Operação"].notna(), "---").fillna("---")
    base["SOA_Origem"] = base.get("SOA_Origem", "---")
    base["SOA_Origem"] = base["SOA_Origem"].fillna("---")

    # Consolidado
    base["Consolidado"] = base.apply(consolidar_registro, axis=1)

    # remover PROCESSANDO do CNM (vazio) e Pendentes do SOA (já vazio)
    base = base[(base["CNM_Status"].fillna("") != "") | (base["SOA_Status"].fillna("") != "") | (base["SGP_Status"].fillna("") != "")]

    # Colunas em ordem; mantém "Operação" e "tipo"
    prefer = [
        "Documento","Nome","Consolidado",
        "CNM_Status","SOA_Status","SGP_Status",
        "Operação","tipo","CNM_Origem","SOA_Origem"
    ]
    exist = [c for c in prefer if c in base.columns]
    base = base[exist + [c for c in base.columns if c not in exist]]

    # Preenche vazios com '---' em colunas-chave
    for c in ["Nome","Consolidado","CNM_Status","SOA_Status","SGP_Status","Operação","tipo","CNM_Origem","SOA_Origem"]:
        if c in base.columns:
            base[c] = base[c].fillna("---").replace("", "---")

    return base


# =========================
# ======= HTML OUT =========
# =========================
def gerar_html(df: pd.DataFrame, out_path: str):
    info(f"Gerando HTML em: {out_path}")
    df2 = df.copy()

    # cria dropdown de filtro por Consolidado e subtítulo dinâmico
    html = f"""
<!doctype html>
<html lang="pt-BR">
<head>
<meta charset="utf-8" />
<title>Dashboard Unificado</title>
<meta name="viewport" content="width=device-width, initial-scale=1" />
<link rel="preconnect" href="https://cdn.jsdelivr.net" />
<link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/datatables.net-dt/css/jquery.dataTables.min.css" />
<style>
body {{ font-family: system-ui, -apple-system, Segoe UI, Roboto, Arial, sans-serif; margin: 16px; }}
h1 {{ margin-bottom: 0; }}
#subtitle {{ margin: 4px 0 16px; color: #444; }}
.toolbar {{ display:flex; gap:12px; align-items:center; margin: 12px 0; flex-wrap: wrap; }}
table.dataTable thead th {{ white-space: nowrap; }}
.badge {{ display:inline-block; padding: 2px 8px; border-radius: 8px; background:#eee; margin-left:6px;}}
</style>
</head>
<body>
  <h1>Dashboard Unificado <span class="badge">{len(df2)} linhas</span></h1>
  <div id="subtitle">Mostrando todos os status.</div>

  <div class="toolbar">
    <label for="statusFilter"><strong>Filtrar por Status:</strong></label>
    <select id="statusFilter">
      <option value="">(Todos)</option>
      <option>NEGATIVADO</option>
      <option>BAIXADO</option>
      <option>ERRO</option>
      <option>---</option>
    </select>
  </div>

  <table id="tbl" class="display" style="width:100%"></table>

  <script src="https://cdn.jsdelivr.net/npm/jquery@3.7.1/dist/jquery.min.js"></script>
  <script src="https://cdn.jsdelivr.net/npm/datatables.net@2.0.8/js/dataTables.min.js"></script>
  <script>
    const columns = {list(df2.columns)};
    const data = {df2.values.tolist()};
    let dt;

    function subtitulo(status) {{
      const el = document.getElementById('subtitle');
      if (!status) {{
        el.textContent = "Mostrando todos os status.";
        return;
      }}
      const frases = {{
        "NEGATIVADO": "Itens ativos em negativação (CNM/SOA/SGP).",
        "BAIXADO": "Registros baixados/regularizados.",
        "ERRO": "Divergências entre origens ou inconsistências mapeadas.",
        "---": "Registros sem determinação consolidada."
      }};
      el.textContent = frases[status] || ("Filtrando por status: " + status);
    }}

    $(function() {{
      dt = new DataTable('#tbl', {{
        data: data,
        columns: columns.map(c => ({{ title: c }})),
        pageLength: 25,
        deferRender: true,
        order: [[0, 'asc']]
      }});

      $('#statusFilter').on('change', function() {{
        const v = this.value;
        const colIdx = columns.indexOf('Consolidado');
        if (colIdx >= 0) {{
          dt.column(colIdx).search(v ? '^' + v + '$' : '', true, false).draw();
        }}
        subtitulo(v);
      }});
    }});
  </script>
</body>
</html>
    """.strip()

    Path(out_path).write_text(html, encoding="utf-8")


# =========================
# ========= MAIN ==========
# =========================
def gerar_dashboard():
    info("Lendo arquivos CNM e SGP...")
    df_cnm = ler_cnm(CNM_FILE)  # preserva Operação & tipo
    df_sgp = ler_sgp(SGP_FILE)

    info("Lendo arquivos SOA (CSV)...")
    df_soa = ler_soa(SOA_CSVS)

    info(f"CNM: linhas={len(df_cnm)}")
    info(f"SGP: linhas={len(df_sgp)}")
    info(f"SOA: linhas={len(df_soa)}")

    info("Construindo unificado...")
    df_all = construir_unificado(df_cnm, df_soa, df_sgp)

    # limpa output e grava
    ensure_output_clean()
    out_html = os.path.join(OUTPUT_DIR, "dashboard_unificado.html")
    out_csv  = os.path.join(OUTPUT_DIR, "dashboard_unificado.csv")

    df_all.to_csv(out_csv, index=False, encoding="utf-8")
    gerar_html(df_all, out_html)

    info(f"HTML gerado em: {out_html}")
    info("Concluído.")


if __name__ == "__main__":
    try:
        gerar_dashboard()
    except Exception as e:
        err(f"Falha na execução: {e}")
        raise
