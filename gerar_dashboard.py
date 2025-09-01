import pandas as pd
import os
import html
import re
import unicodedata
from collections import Counter

# === CONFIGURAÇÃO DE CAMINHOS ===
caminho_dir = './download'
saida_dir = './output'
os.makedirs(saida_dir, exist_ok=True)

caminho_cnm = os.path.join(caminho_dir, 'Relatorio_CNM.xlsx')
caminho_sgp = os.path.join(caminho_dir, 'Relatorio_SGP.xlsx')

arquivos_soa = {
    "Ativas": "Ativas.csv",
    "Baixadas": "Baixadas.csv",
    "Pendentes": "Pendentes.csv",
    "Determinacao": "Determinacao.csv",
    "Erros": "Erros.csv"
}

# ==========================
# Helpers de saída
# ==========================
def limpar_output():
    """Remove todos os arquivos existentes na pasta de saída antes de gerar novos."""
    if not os.path.isdir(saida_dir):
        return
    for nome in os.listdir(saida_dir):
        caminho = os.path.join(saida_dir, nome)
        try:
            if os.path.isfile(caminho):
                os.remove(caminho)
        except Exception:
            pass

# ==========================
# Helpers de colunas/cabeçalho/normalização
# ==========================
def deduplicar_colunas(cols):
    counter = Counter()
    novas = []
    for col in cols:
        counter[col] += 1
        novas.append(col if counter[col] == 1 else f"{col}.{counter[col]-1}")
    return novas

def _norm(s: str) -> str:
    if s is None:
        return ""
    s = str(s).strip()
    s = unicodedata.normalize("NFKD", s)
    s = "".join(ch for ch in s if not unicodedata.combining(ch))
    s = re.sub(r"[\s_/|-]+", " ", s)
    s = re.sub(r"[^0-9A-Za-z ]", "", s)
    return s.upper().strip()

def _upper_no_accents(s: str) -> str:
    if s is None:
        return ""
    s = unicodedata.normalize("NFKD", str(s).strip())
    s = "".join(ch for ch in s if not unicodedata.combining(ch))
    return s.upper().strip()

def encontrar_coluna(df, candidatos):
    """Localiza primeira coluna cujo nome (normalizado) casa exatamente ou contém o candidato normalizado."""
    mapa_norm = {col: _norm(col) for col in df.columns}
    for cand in candidatos:
        cand_norm = _norm(cand)
        # match exato
        for col, normv in mapa_norm.items():
            if normv == cand_norm:
                return col
        # match por inclusão
        for col, normv in mapa_norm.items():
            if cand_norm in normv:
                return col
    return None

def renomear_colunas_para_padrao(df, origem="CNM"):
    """
    Padroniza nomes-chave:
      - 'CPF/CNPJ'
      - (CNM) 'Tipo', 'Data / Hora' (somente rótulos de operação), 'Nome'
    """
    mapeamento_comum = {
        "CPF/CNPJ": [
            "CPF/CNPJ", "CPF - CNPJ", "CPF CNPJ", "Documento", "DOC", "CPF", "CNPJ", "Documento (CPF)"
        ]
    }
    mapeamento_cnm_extra = {
        "Tipo": [
            "Tipo", "Tipo Operacao", "Tipo Operação", "Operacao", "Operação", "Evento", "Movimento", "Tipo Movimento"
        ],
        # ⚠️ Evita 'Data' genérica; usa apenas rótulos ligados à operação
        "Data / Hora": [
            "Data / Hora", "Data Hora", "Data Operacao", "Data Operação",
            "Data Inclusao", "Data Inclusão", "Data Exclusao", "Data Exclusão"
        ],
        "Nome": [
            "Nome", "Devedor", "Razao Social", "Razão Social", "Cliente", "Pessoa", "Nome/Razão Social", "Nome Razao Social"
        ]
    }

    ren = {}
    for alvo, cand in mapeamento_comum.items():
        col = encontrar_coluna(df, cand)
        if col: ren[col] = alvo

    if origem.upper() == "CNM":
        for alvo, cand in mapeamento_cnm_extra.items():
            col = encontrar_coluna(df, cand)
            if col: ren[col] = alvo

    return df.rename(columns=ren) if ren else df

# ==========================
# Leitura inteligente de Excel (detecta header real)
# ==========================
def _escolher_melhor_aba(xl_dict):
    melhor_nome, melhor_score = None, -1
    for nome, df in xl_dict.items():
        score = df.notna().sum().sum()
        if score > melhor_score:
            melhor_nome, melhor_score = nome, score
    return melhor_nome

def _detectar_linha_cabecalho(df, candidatos_header=None, max_scan=30):
    if candidatos_header is None:
        candidatos_header = ["CPF/CNPJ", "CPF - CNPJ", "CPF", "CNPJ", "Documento", "Tipo", "Data / Hora", "Nome", "Devedor"]
    cand_norm = [_norm(c) for c in candidatos_header]
    n_scan = min(len(df), max_scan)
    for i in range(n_scan):
        linha = df.iloc[i].astype(str).fillna("").tolist()
        linha_norm = [_norm(x) for x in linha]
        hits = sum(1 for val in linha_norm if any(c in val for c in cand_norm))
        if hits >= 2:
            return i
    return 0

def ler_excel_com_header_real(caminho, origem="CNM", skiprows_sugerido=None):
    xls = pd.read_excel(caminho, sheet_name=None, header=None, dtype=str, engine='openpyxl')
    aba = _escolher_melhor_aba(xls)
    df = xls[aba].copy()

    if skiprows_sugerido:
        df = df.iloc[skiprows_sugerido:].reset_index(drop=True)

    header_idx = _detectar_linha_cabecalho(df)
    header = df.iloc[header_idx].astype(str).fillna("").tolist()
    header = [h if not re.match(r"^Unnamed", str(h), re.I) else "" for h in header]
    header = [h.strip() for h in header]

    if not any(h for h in header) and header_idx + 1 < len(df):
        header_idx += 1
        header = df.iloc[header_idx].astype(str).fillna("").tolist()
        header = [h if not re.match(r"^Unnamed", str(h), re.I) else "" for h in header]
        header = [h.strip() for h in header]

    header = [h if h else f"COL_{i}" for i, h in enumerate(header)]
    if len(set(header)) < len(header):
        header = deduplicar_colunas(header)

    df = df.iloc[header_idx + 1:].reset_index(drop=True)
    df.columns = header
    df = df.dropna(axis=1, how='all').dropna(axis=0, how='all')
    df = df.loc[:, ~df.columns.str.match(r"^Unnamed", na=False)]

    return renomear_colunas_para_padrao(df, origem=origem)

# ==========================
# Documento & Data
# ==========================
def padronizar_cpf_cnpj(coluna):
    """Normaliza CPF (11) e CNPJ (14)."""
    def _n(v):
        d = re.sub(r'\D', '', str(v))
        if len(d) <= 11:
            return d[-11:].zfill(11)
        else:
            return d[-14:].zfill(14)
    return coluna.astype(str).map(_n)

def to_datetime_robusta(series):
    a = pd.to_datetime(series, errors='coerce', dayfirst=True)
    b = pd.to_datetime(series, errors='coerce', dayfirst=False)
    return a.fillna(b)

def parse_data_hora(valor):
    s = str(valor).strip()
    if not s or s in {"NaT", "nan"}:
        return "-"
    try:
        if re.match(r"^\d{4}-\d{2}-\d{2}", s):
            dt = pd.to_datetime(s, dayfirst=False, errors='coerce')
        else:
            dt = pd.to_datetime(s, dayfirst=True, errors='coerce')
        return s if pd.isna(dt) else dt.strftime('%d/%m/%Y %H:%M')
    except Exception:
        return s

# ==========================
# SOA
# ==========================
def ler_e_normalizar_soa(nome, caminho_arquivo):
    print(f"[INFO] Lendo arquivo SOA: {nome} -> {caminho_arquivo}")
    # Sniff automático de separador (',' ou ';')
    df = pd.read_csv(caminho_arquivo, encoding="utf-8", dtype=str, sep=None, engine='python')
    raw_cols = list(df.columns)
    cols = [html.unescape(col).replace('+ACI-', '').replace('+AC0-', '-').replace('"', '').strip() for col in raw_cols]
    if len(set(cols)) < len(cols):
        print(f"[AVISO] Colunas duplicadas detectadas em {nome}. Renomeando...")
        cols = deduplicar_colunas(cols)
    df.columns = cols

    df = df.rename(columns={
        'Documento': 'documento',
        'Devedor': 'devedor',
        'Unique ID': 'Unique ID',
        'Data Inclusão': 'data',
        'Data Inclusao': 'data',
        'Data Exclusão': 'data',
        'Data Exclusao': 'data'
    })
    df = df.loc[:, ~df.columns.duplicated()]

    possivel_doc = 'documento' if 'documento' in df.columns else encontrar_coluna(df, ["Documento", "CPF/CNPJ", "CPF - CNPJ", "CPF", "CNPJ"])
    if possivel_doc:
        df['CPF/CNPJ'] = padronizar_cpf_cnpj(df[possivel_doc])
    else:
        df['CPF/CNPJ'] = ""

    df['fonte'] = nome
    return df

# ==========================
# Carga de dados
# ==========================
def carregar_dados():
    if not os.path.exists(caminho_cnm) or not os.path.exists(caminho_sgp):
        print('[ERRO] Arquivo CNM ou SGP não encontrado.')
        exit(1)

    print('[INFO] Lendo arquivos CNM e SGP...')
    cnm_df = ler_excel_com_header_real(caminho_cnm, origem="CNM")

    try:
        sgp_df = ler_excel_com_header_real(caminho_sgp, origem="SGP", skiprows_sugerido=8)
    except Exception:
        sgp_df = ler_excel_com_header_real(caminho_sgp, origem="SGP", skiprows_sugerido=None)

    if 'CPF/CNPJ' not in cnm_df.columns:
        raise KeyError(f"[CNM] Coluna 'CPF/CNPJ' não encontrada. Colunas: {list(cnm_df.columns)}")
    if 'CPF/CNPJ' not in sgp_df.columns:
        raise KeyError(f"[SGP] Coluna 'CPF/CNPJ' não encontrada. Colunas: {list(sgp_df.columns)}")

    cnm_df['CPF/CNPJ'] = padronizar_cpf_cnpj(cnm_df['CPF/CNPJ'])
    sgp_df['CPF/CNPJ'] = padronizar_cpf_cnpj(sgp_df['CPF/CNPJ'])

    if 'Tipo' not in cnm_df.columns:
        cnm_df['Tipo'] = "ERRO"

    # Data de operação apenas por rótulos claros
    if 'Data / Hora' in cnm_df.columns:
        cnm_df['_dt'] = to_datetime_robusta(cnm_df['Data / Hora'])
    else:
        cnm_df['_dt'] = pd.NaT

    # SOA
    lista_soa = []
    for nome, arquivo in arquivos_soa.items():
        caminho = os.path.join(caminho_dir, arquivo)
        if os.path.exists(caminho):
            lista_soa.append(ler_e_normalizar_soa(nome, caminho))
        else:
            print(f"[AVISO] Arquivo {arquivo} não encontrado.")
    if lista_soa:
        soa_df = pd.concat(lista_soa, ignore_index=True)
        for col in ["CPF/CNPJ", "devedor", "data", "fonte", "Unique ID"]:
            if col not in soa_df.columns:
                soa_df[col] = "-"
    else:
        soa_df = pd.DataFrame(columns=["CPF/CNPJ", "devedor", "data", "fonte", "Unique ID"])

    return cnm_df, sgp_df, soa_df

# ==========================
# Determinar evento do CNM (mais recente) por CPF/CNPJ IGNORANDO PF/PJ
# ==========================
def _classificar_evento(valor_tipo: str):
    """
    Classifica INCLUSAO / EXCLUSAO por substring em valor normalizado (sem acento).
    Retorna "INCLUSAO", "EXCLUSAO" ou None.
    """
    t = _upper_no_accents(valor_tipo)
    if not t:
        return None
    if "EXCLUSAO" in t:
        return "EXCLUSAO"
    if "INCLUSAO" in t:
        return "INCLUSAO"
    return None  # PF/PJ/consultas e demais

def obter_tipo_cnm_mais_recente(rows_cnm: pd.DataFrame):
    """
    Considera apenas INCLUSAO/EXCLUSAO (por substring), ignorando PF/PJ.
    - Se houver datas válidas (_dt) → pega a mais RECENTE entre os relevantes.
    - Sem datas → prioridade EXCLUSAO > INCLUSAO; senão "-"
    """
    if rows_cnm.empty or 'Tipo' not in rows_cnm.columns:
        return "-"

    classificados = rows_cnm.assign(_ev=rows_cnm['Tipo'].map(_classificar_evento))
    relevantes = classificados[~classificados['_ev'].isna()].copy()
    if relevantes.empty:
        return "-"

    if '_dt' in relevantes.columns and relevantes['_dt'].notna().any():
        idx = relevantes['_dt'].idxmax()
        return relevantes.loc[idx, '_ev']

    # Sem data: prioridade EXCLUSAO > INCLUSAO
    if (relevantes['_ev'] == "EXCLUSAO").any():
        return "EXCLUSAO"
    if (relevantes['_ev'] == "INCLUSAO").any():
        return "INCLUSAO"
    return "-"

def obter_data_cnm_mais_recente(rows_cnm: pd.DataFrame):
    if rows_cnm.empty:
        return "-"
    if '_dt' in rows_cnm.columns and rows_cnm['_dt'].notna().any():
        dt = rows_cnm.loc[rows_cnm['_dt'].idxmax(), '_dt']
        try:
            return dt.strftime('%d/%m/%Y %H:%M')
        except Exception:
            pass
    if 'Data / Hora' in rows_cnm.columns:
        for v in rows_cnm['Data / Hora']:
            s = parse_data_hora(v)
            if s != "-":
                return s
    return "-"

# ==========================
# Regras finais de TIPO + STATUS (único ponto que escreve esses campos)
# ==========================
def definir_tipo_status(tipo_cnm_final, doc, documentos_sgp, fontes_soa, tem_cnm):
    """
    Regras:
    - Se TEM CNM:
        - Se DOC está no SGP:
            - CNM = EXCLUSAO  -> TIPO = Baixada,   STATUS = ERRO
            - Caso contrário  -> TIPO = NEGATIVADO, STATUS = NEGATIVADO
              (inclui CNM=INCLUSAO ou CNM não classificável/indefinido)
        - Se DOC NÃO está no SGP:
            - CNM = EXCLUSAO  -> TIPO = Baixada,   STATUS = BAIXADO
            - CNM = INCLUSAO  -> TIPO = NEGATIVADO, STATUS = ERRO
            - Caso contrário  -> TIPO = ERRO,      STATUS = ERRO
    - Se NÃO TEM CNM (fallback via SOA):
        - ATIVAS        -> NEGATIVADO / NEGATIVADO
        - BAIXADAS      -> Baixada   / BAIXADO
        - DETERMINACAO  -> Determinação / ERRO
        - ERROS/PENDENTES -> ERRO / ERRO
        - Sem SOA → ERRO / ERRO
    """
    if tem_cnm:
        if doc in documentos_sgp:
            if tipo_cnm_final == "EXCLUSAO":
                return ("Baixada", "ERRO")
            return ("NEGATIVADO", "NEGATIVADO")
        else:
            if tipo_cnm_final == "EXCLUSAO":
                return ("Baixada", "BAIXADO")
            if tipo_cnm_final == "INCLUSAO":
                return ("NEGATIVADO", "ERRO")
            return ("ERRO", "ERRO")
    else:
        fontes_upper = {_upper_no_accents(f) for f in (fontes_soa or [])}
        if "ATIVAS" in fontes_upper:
            return ("NEGATIVADO", "NEGATIVADO")
        if "BAIXADAS" in fontes_upper:
            return ("Baixada", "BAIXADO")
        if "DETERMINACAO" in fontes_upper:
            return ("Determinação", "ERRO")
        if "ERROS" in fontes_upper or "PENDENTES" in fontes_upper:
            return ("ERRO", "ERRO")
        return ("ERRO", "ERRO")


# ==========================
# Geração principal (único lugar que define e escreve TIPO/STATUS)
# ==========================
def gerar_dashboard():
    # 1) Limpa a pasta de saída
    limpar_output()

    # 2) Carrega bases
    cnm_df, sgp_df, soa_df = carregar_dados()

    documentos = set(cnm_df['CPF/CNPJ']) | set(soa_df['CPF/CNPJ']) | set(sgp_df['CPF/CNPJ'])
    documentos_sgp = set(sgp_df['CPF/CNPJ'])

    dados = []
    for doc in documentos:
        row_cnm = cnm_df[cnm_df['CPF/CNPJ'] == doc]
        row_sgp = sgp_df[sgp_df['CPF/CNPJ'] == doc]
        row_soa = soa_df[soa_df['CPF/CNPJ'] == doc]

        tem_cnm = not row_cnm.empty
        fontes = list(row_soa['fonte'].unique()) if not row_soa.empty else []

        # Evento CNM/Data MAIS RECENTES (ignorando PF/PJ e aceitando substrings)
        if tem_cnm:
            tipo_cnm_final = obter_tipo_cnm_mais_recente(row_cnm)
            data = obter_data_cnm_mais_recente(row_cnm)
        elif not row_soa.empty:
            tipo_cnm_final = "-"
            data = str(row_soa['data'].values[0]) if 'data' in row_soa.columns else "-"
        else:
            tipo_cnm_final = "-"
            data = "-"

        # === *** ÚNICO ponto que escreve TIPO e STATUS *** ===
        tipo_norm, status = definir_tipo_status(tipo_cnm_final, doc, documentos_sgp, fontes, tem_cnm)

        # Nome / ID
        nome = row_soa['devedor'].values[0] if not row_soa.empty else (row_cnm['Nome'].values[0] if ('Nome' in row_cnm.columns and tem_cnm and not row_cnm['Nome'].isna().all()) else "-")
        id_val = row_soa['Unique ID'].values[0] if ('Unique ID' in row_soa.columns and not row_soa['Unique ID'].isnull().all() and not row_soa.empty) else "-"

        # Local (verde/vermelho)
        locais = []
        for origem, df_rows in [('CNM', row_cnm), ('SOA', row_soa), ('SGP', row_sgp)]:
            color = "green" if not df_rows.empty else "red"
            locais.append(f"<span style='color:{color}'>{origem}</span>")
        local = " | ".join(locais)


        dados.append({
            "ID": str(id_val),
            "CPF/CNPJ": str(doc),
            "Nome": nome,
            "Data": data,
            "Tipo": tipo_norm,   # ← gravando TIPO final aqui
            "Local": local,
            "Status": status     # ← gravando STATUS final aqui
        })

    df = pd.DataFrame(dados)
    df["CPF/CNPJ"] = df["CPF/CNPJ"].astype(str)
    df["ID"] = df["ID"].astype(str)

    # Exporta Excel
    df.to_excel(os.path.join(saida_dir, "resultado_unificado.xlsx"), index=False)
    print("[SUCESSO] resultado_unificado.xlsx gerado com sucesso!")

    gerar_html(df)

# ==========================
# HTML
# ==========================
def gerar_html(df):
    html_tabela = df.to_html(index=False, escape=False, table_id='tabela', classes='display')
    subtitulos = {
        "NEGATIVADO": "Clientes negativados no CNM ou SOA e presentes no SGP.",
        "BAIXADO": "Clientes excluídos no CNM ou no SOA e ausentes no SGP.",
        "ERRO": "Clientes com inconsistência entre CNM, SOA e SGP.",
        "SCORE": "Consultas de SCORE identificadas como PF ou PJ."
    }
    html_code = f"""<!DOCTYPE html>
<html lang='pt-BR'>
<head>
    <meta charset='UTF-8'><title>Dashboard Unificado</title>
    <link rel='stylesheet' href='https://cdn.datatables.net/1.13.4/css/jquery.dataTables.min.css'>
    <script src='https://code.jquery.com/jquery-3.7.0.min.js'></script>
    <script src='https://cdn.datatables.net/1.13.4/js/jquery.dataTables.min.js'></script>
    <style>
        body {{ font-family: Arial, sans-serif; text-align: center; }}
        table {{ margin: 0 auto; width: 90%; }}
        .subtitulo {{ margin: 10px 0; font-weight: bold; }}
        th {{ white-space: nowrap; }}
    </style>
</head>
<body>
    <h2>Dashboard Unificado - CNM + SOA x SGP</h2>
    <div class='subtitulo' id='subtitulo'></div>
    {html_tabela}
    <footer>VERDE = Presente | VERMELHO = Ausente</footer>
    <script>
        const subtitulos = {{
            "NEGATIVADO": "{subtitulos['NEGATIVADO']}",
            "BAIXADO": "{subtitulos['BAIXADO']}",
            "ERRO": "{subtitulos['ERRO']}",
            "SCORE": "{subtitulos['SCORE']}"
        }};
        $(document).ready(function() {{
            const table = $('#tabela').DataTable({{
                paging: true,
                searching: true,
                ordering: true,
                pageLength: 25,
                language: {{
                    url: '//cdn.datatables.net/plug-ins/1.13.4/i18n/pt-BR.json'
                }},
                initComplete: function () {{
                    // 0=ID, 1=CPF/CNPJ, 2=Nome, 3=Data, 4=Tipo, 5=Local, 6=Status
                    const colunasParaFiltrar = {{
                        3: "Data",
                        4: "Tipo",
                        5: "Local",
                        6: "Status"
                    }};
                    this.api().columns().every(function (index) {{
                        if (colunasParaFiltrar[index]) {{
                            const column = this;
                            const label = colunasParaFiltrar[index];
                            const select = $('<select><option value="">Filtrar por ' + label + '</option></select>')
                                .appendTo($(column.header()).empty())
                                .on('change', function () {{
                                    const val = $.fn.dataTable.util.escapeRegex($(this).val());
                                    column.search(val ? '^' + val + '$' : '', true, false).draw();
                                }});
                            column.data().unique().sort().each(function (d) {{
                                if (d) select.append('<option value="' + d + '">' + d + '</option>');
                            }});
                        }}
                    }});
                }}
            }});
            table.on('draw search.dt', function () {{
                const statusColData = table.column(6, {{search: 'applied'}}).data().toArray();
                const unicos = [...new Set(statusColData.filter(Boolean))];
                const msg = unicos.length === 1 ? subtitulos[unicos[0]] || '' : '';
                $('#subtitulo').html(msg);
            }});
        }});
    </script>
</body>
</html>"""
    with open(os.path.join(saida_dir, "dashboard_unificado.html"), "w", encoding="utf-8") as f:
        f.write(html_code)
    print("[SUCESSO] dashboard_unificado.html gerado com sucesso!")

# ==========================
# Main
# ==========================
if __name__ == '__main__':
    gerar_dashboard()
