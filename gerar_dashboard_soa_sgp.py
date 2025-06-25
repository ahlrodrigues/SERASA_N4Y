import pandas as pd
import os
import html
from collections import Counter

# Caminhos dos arquivos
caminho_dir = './download'
saida_dir = './output'
os.makedirs(saida_dir, exist_ok=True)

caminho_sgp = os.path.join(caminho_dir, 'Relatorio_SGP.xlsx')

arquivos_soa = {
    "Ativas": "Ativas.csv",
    "Baixadas": "Baixadas.csv",
    "Pendentes": "Pendentes.csv",
    "Determinacao": "Determinacao.csv",
    "Erros": "Erros.csv"
}

def deduplicar_colunas(cols):
    counter = Counter()
    novas = []
    for col in cols:
        counter[col] += 1
        if counter[col] == 1:
            novas.append(col)
        else:
            novas.append(f"{col}.{counter[col]-1}")
    return novas

def normalizar_documento(coluna):
    return (
        coluna.astype(str)
            .str.replace(r'\.0$', '', regex=True)
            .str.replace(r'\D', '', regex=True)
            .str.strip()
    )

def ler_e_normalizar_soa(nome, caminho_arquivo):
    print(f"[INFO] Lendo arquivo: {nome} -> {caminho_arquivo}")
    df = pd.read_csv(caminho_arquivo, encoding="utf-8")
    raw_cols = list(df.columns)
    cols = [html.unescape(col).replace('+ACI-', '').replace('+AC0-', '-').replace('"', '').strip() for col in raw_cols]
    if len(set(cols)) < len(cols):
        print(f"[AVISO] Colunas duplicadas detectadas em {nome}. Renomeando com sufixos...")
        cols = deduplicar_colunas(cols)
    df.columns = cols
    df = df.rename(columns={
        'Documento': 'documento',
        'Devedor': 'devedor',
        'Unique ID': 'Unique ID',
        'Data Inclusão': 'data',
        'Data Inclusao': 'data',
        'Data Exclusão': 'data'
    })
    df = df.loc[:, ~df.columns.duplicated()]
    df['documento'] = normalizar_documento(df['documento'])
    df['fonte'] = nome
    return df

def carregar_dados():
    if not os.path.exists(caminho_sgp):
        print('[ERRO] Arquivo SGP não encontrado.')
        exit(1)

    print('[INFO] Lendo arquivo SGP...')
    sgp_df = pd.read_excel(caminho_sgp, skiprows=8)
    sgp_df['CPF/CNPJ'] = normalizar_documento(sgp_df['CPF/CNPJ'])

    documentos_sgp = set(sgp_df['CPF/CNPJ'])

    lista_soa = []
    for nome, arquivo in arquivos_soa.items():
        caminho = os.path.join(caminho_dir, arquivo)
        if os.path.exists(caminho):
            df = ler_e_normalizar_soa(nome, caminho)
            lista_soa.append(df)
        else:
            print(f"[AVISO] Arquivo {arquivo} não encontrado.")

    soa_df = pd.concat(lista_soa, ignore_index=True) if lista_soa else pd.DataFrame(columns=["documento", "devedor", "data", "fonte", "Unique ID"])
    return soa_df, sgp_df

def determinar_status(presente_sgp, tipo):
    tipo = tipo.strip().upper()
    if tipo == "INCLUSAO":
        return "NEGATIVADO" if presente_sgp else "ERRO"
    elif tipo == "EXCLUSAO":
        return "ERRO" if presente_sgp else "BAIXADO"
    else:
        return "ERRO"

def gerar_dashboard():
    soa_df, sgp_df = carregar_dados()
    documentos = set(soa_df['documento']) | set(sgp_df['CPF/CNPJ'])

    dados = []
    for doc in documentos:
        row_soa = soa_df[soa_df['documento'] == doc]
        row_sgp = sgp_df[sgp_df['CPF/CNPJ'] == doc]
        presente_sgp = not row_sgp.empty

        if not row_soa.empty:
            tipo = row_soa['fonte'].values[0]
            data = row_soa['data'].values[0]
            nome = row_soa['devedor'].values[0]
            id_val = row_soa['Unique ID'].values[0] if 'Unique ID' in row_soa and not row_soa['Unique ID'].isnull().all() else "-"
        else:
            tipo = "-"
            data = "-"
            nome = "-"
            id_val = "-"

        local = " | ".join([
            f"<span style='color:{'green' if not row.empty else 'red'}'>{origem}</span>"
            for origem, row in [('SOA', row_soa), ('SGP', row_sgp)]
        ])

        status = determinar_status(presente_sgp, tipo)

        dados.append({
            "ID": id_val,
            "Documento": doc,
            "Nome": nome,
            "Data": data,
            "Tipo": tipo,
            "Local": local,
            "Status": status
        })

    df = pd.DataFrame(dados)
    df.to_excel(os.path.join(saida_dir, "resultado_soa_sgp.xlsx"), index=False)
    gerar_html(df)

def gerar_html(df):
    html_tabela = df.to_html(index=False, escape=False, table_id='tabela', classes='display')
    subtitulos = {
        "NEGATIVADO": "Clientes negativados no SOA e presentes no SGP.",
        "BAIXADO": "Clientes excluídos no SOA e ausentes no SGP.",
        "ERRO": "Clientes com inconsistência entre SOA e SGP."
    }
    html_code = f"""<!DOCTYPE html>
<html lang='pt-BR'>
<head>
    <meta charset='UTF-8'><title>Dashboard SOA x SGP</title>
    <link rel='stylesheet' href='https://cdn.datatables.net/1.13.4/css/jquery.dataTables.min.css'>
    <script src='https://code.jquery.com/jquery-3.7.0.min.js'></script>
    <script src='https://cdn.datatables.net/1.13.4/js/jquery.dataTables.min.js'></script>
    <style>
        body {{ font-family: Arial; text-align: center; }}
        table {{ margin: 0 auto; width: 90%; }}
        .subtitulo {{ margin: 10px 0; font-weight: bold; }}
    </style>
</head>
<body>
    <h2>Dashboard SOA x SGP</h2>
    <div class='subtitulo' id='subtitulo'></div>
    {html_tabela}
    <footer>VERDE = Presente | VERMELHO = Ausente</footer>
    <script>
        const subtitulos = {subtitulos};
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
                    this.api().columns().every(function (index) {{
                        if (index === 3 || index === 4 || index === 6) {{
                            const column = this;
                            const select = $('<select><option value=""></option></select>')
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
            table.on('search.dt', function () {{
                const val = table.column(6, {{search: 'applied'}}).data()[0];
                $('#subtitulo').html(subtitulos[val] || '');
            }});
        }});
    </script>
</body>
</html>"""
    with open(os.path.join(saida_dir, "dashboard_soa_sgp.html"), "w", encoding="utf-8") as f:
        f.write(html_code)

if __name__ == '__main__':
    gerar_dashboard()
