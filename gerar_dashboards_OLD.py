
import pandas as pd
import os
import zipfile
import html

# === Caminhos dos arquivos ===
caminho_cnm = './download/Relatorio_CNM.xlsx'
caminho_sgp = './download/Relatorio_SGP.xlsx'
caminho_soa = './download/Ativas.csv'
saida_dir = './output'
os.makedirs(saida_dir, exist_ok=True)

def validar_arquivo_excel(caminho):
    if not os.path.exists(caminho):
        raise FileNotFoundError(f"[ERRO] Arquivo não encontrado: {caminho}")
    if os.path.getsize(caminho) < 10240:
        raise ValueError(f"[ERRO] Arquivo muito pequeno: {caminho}")
    try:
        with zipfile.ZipFile(caminho, 'r') as zf:
            if zf.testzip() is not None:
                raise zipfile.BadZipFile("Conteúdo ZIP inválido")
    except zipfile.BadZipFile as e:
        raise zipfile.BadZipFile(f"[ERRO] Arquivo .xlsx corrompido: {e}")

def normalizar_documento(coluna):
    return (
        coluna.astype(str)
              .str.replace(r'\.0$', '', regex=True)
              .str.replace(r'\D', '', regex=True)
              .str.strip()
    )

def comparar_cnm_sgp():
    validar_arquivo_excel(caminho_cnm)
    validar_arquivo_excel(caminho_sgp)
    cnm_df = pd.read_excel(caminho_cnm, engine="openpyxl")
    sgp_df = pd.read_excel(caminho_sgp, skiprows=8)
    cnm_df['Documento'] = normalizar_documento(cnm_df['Documento'])
    sgp_df['CPF/CNPJ'] = normalizar_documento(sgp_df['CPF/CNPJ'])
    todos_documentos = set(cnm_df['Documento']).union(set(sgp_df['CPF/CNPJ']))
    dados = []
    for doc in todos_documentos:
        row_cnm = cnm_df[cnm_df['Documento'] == doc]
        row_sgp = sgp_df[sgp_df['CPF/CNPJ'] == doc]
        presente_cnm = not row_cnm.empty
        presente_sgp = not row_sgp.empty
        tipo = row_cnm['Tipo'].values[0] if presente_cnm else ""
        nome = row_sgp['Nome/Razão Social'].values[0] if presente_sgp else ""
        id_val = int(row_cnm['Id'].values[0]) if presente_cnm else ""
        data = pd.to_datetime(row_cnm['Data / Hora'].values[0], dayfirst=True).strftime('%d/%m/%Y %H:%M') if presente_cnm else ""
        local = f"<span style='color:{'green' if presente_cnm else 'red'}'>CNM</span> | <span style='color:{'green' if presente_sgp else 'red'}'>SGP</span>"
        if presente_cnm and tipo == 'INCLUSAO' and presente_sgp:
            status = 'NEGATIVADO'
        elif presente_cnm and tipo == 'EXCLUSAO' and not presente_sgp:
            status = 'BAIXADO'
        else:
            status = 'ERRO'
        dados.append({ "Id": id_val, "Documento": doc, "Nome": nome, "Data": data, "Tipo": tipo, "Local": local, "Status": status })
    df = pd.DataFrame(dados)
    df.to_excel(os.path.join(saida_dir, 'resultado_cnm_sgp.xlsx'), index=False)
    gerar_html(df, 'dashboard_cnm_sgp.html', 'Dashboard CNM x SGP')
    return df

def comparar_soa_sgp():
    validar_arquivo_excel(caminho_sgp)
    raw_df = pd.read_csv(caminho_soa, encoding="utf-8")
    decoded_columns = [html.unescape(col).replace('+ACI-', '').replace('+AC0-', '-').replace('"', '').strip() for col in raw_df.columns]
    raw_df.columns = decoded_columns
    soa_df = raw_df.rename(columns={ '+//3//f/9-Unique ID': 'Unique ID', 'Documento': 'documento', 'Devedor': 'devedor', 'Data Inclus+//3//Q-o': 'Data Inclusão' })
    soa_df['documento'] = normalizar_documento(soa_df['documento'])
    soa_df['Unique ID'] = soa_df['Unique ID'].astype(str).apply(html.unescape).str.replace(r'\W', '', regex=True)
    soa_df['devedor'] = soa_df['devedor'].astype(str).apply(html.unescape)
    soa_df['Data Inclusão'] = soa_df['Data Inclusão'].astype(str).apply(html.unescape)
    sgp_df = pd.read_excel(caminho_sgp, skiprows=8)
    sgp_df['CPF/CNPJ'] = normalizar_documento(sgp_df['CPF/CNPJ'])
    docs_soa = set(soa_df['documento'])
    docs_sgp = set(sgp_df['CPF/CNPJ'])
    todos_docs = docs_soa.union(docs_sgp)
    dados = []
    for doc in todos_docs:
        row_soa = soa_df[soa_df['documento'] == doc]
        row_sgp = sgp_df[sgp_df['CPF/CNPJ'] == doc]
        presente_soa = not row_soa.empty
        presente_sgp = not row_sgp.empty
        id_val = row_soa['Unique ID'].values[0] if presente_soa else "-"
        nome = row_soa['devedor'].values[0] if presente_soa else (row_sgp['Nome/Razão Social'].values[0] if presente_sgp else "-")
        data = row_soa['Data Inclusão'].values[0] if presente_soa else "-"
        local = f"<span style='color:{'green' if presente_soa else 'red'}'>SOA</span> | <span style='color:{'green' if presente_sgp else 'red'}'>SGP</span>"
        if presente_soa and presente_sgp:
            tipo = "INCLUSÃO"
            status = "NEGATIVADO"
        else:
            tipo = "ERRO"
            status = "ERRO"
        dados.append({ "ID": id_val, "Documento": doc, "Nome": nome, "Data Inclusão": data, "Local": local, "Tipo": tipo, "Status": status })
    df = pd.DataFrame(dados)
    df.to_excel(os.path.join(saida_dir, 'resultado_soa_sgp.xlsx'), index=False)
    gerar_html(df, 'dashboard_soa_sgp.html', 'Dashboard SOA x SGP')
    return df

def gerar_html(df, nome_arquivo, titulo):
    html_tabela = df.to_html(index=False, escape=False, table_id='tabela', classes='display')
    subtitulos = {
        "NEGATIVADO": "Relação de clientes que estão negativados nas duas planilhas.",
        "BAIXADO": "Relação de clientes que estão baixados nas duas planilhas..",
        "ERRO": "Relação de clientes que não estão presentes em uma das planilhas."
    }
    html_code = f"""<!DOCTYPE html>
<html lang='pt-BR'>
<head>
    <meta charset='UTF-8'><title>{titulo}</title>
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
    <h2>{titulo}</h2>
    <div class='subtitulo' id='subtitulo'></div>
    {html_tabela}
    <footer>VERDE = Presente | VERMELHO = Ausente</footer>
    <script>
        const subtitulos = {{
            "NEGATIVADO": "{subtitulos['NEGATIVADO']}",
            "BAIXADO": "{subtitulos['BAIXADO']}",
            "ERRO": "{subtitulos['ERRO']}"
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
                    const column = this.api().column(-1);
                    const select = $('<select><option value="">Filtrar por Status</option></select>')
                        .appendTo($(column.header()).empty())
                        .on('change', function () {{
                            const val = $.fn.dataTable.util.escapeRegex($(this).val());
                            column.search(val ? '^' + val + '$' : '', true, false).draw();
                        }});
                    column.data().unique().sort().each(function (d) {{
                        select.append('<option value="' + d + '">' + d + '</option>');
                    }});
                }}
            }});
            table.on('search.dt', function () {{
                const val = table.column(-1, {{search: 'applied'}}).data()[0];
                $('#subtitulo').html(subtitulos[val] || '');
            }});
        }});
    </script>
</body>
</html>"""
    with open(os.path.join(saida_dir, nome_arquivo), 'w', encoding='utf-8') as f:
        f.write(html_code)

if __name__ == '__main__':
    comparar_cnm_sgp()
    comparar_soa_sgp()
