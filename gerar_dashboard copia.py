
import pandas as pd
import os
import zipfile
import html

# Caminhos
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

def comparar_tudo():
    validar_arquivo_excel(caminho_cnm)
    validar_arquivo_excel(caminho_sgp)
    cnm_df = pd.read_excel(caminho_cnm, engine='openpyxl')
    sgp_df = pd.read_excel(caminho_sgp, skiprows=8)
    raw_soa = pd.read_csv(caminho_soa, encoding='utf-8')
    decoded_cols = [html.unescape(col).replace('+ACI-', '').replace('+AC0-', '-').replace('"', '').strip() for col in raw_soa.columns]
    raw_soa.columns = decoded_cols
    soa_df = raw_soa.rename(columns={
        '+//3//f/9-Unique ID': 'Unique ID',
        'Documento': 'documento',
        'Devedor': 'devedor',
        'Data Inclus+//3//Q-o': 'Data Inclusão'
    })

    # Normalização
    cnm_df['Documento'] = normalizar_documento(cnm_df['Documento'])
    sgp_df['CPF/CNPJ'] = normalizar_documento(sgp_df['CPF/CNPJ'])
    soa_df['documento'] = normalizar_documento(soa_df['documento'])
    soa_df['Unique ID'] = soa_df['Unique ID'].astype(str).apply(html.unescape).str.replace(r'\W', '', regex=True)
    soa_df['devedor'] = soa_df['devedor'].astype(str).apply(html.unescape)
    soa_df['Data Inclusão'] = soa_df['Data Inclusão'].astype(str).apply(html.unescape)

    # União dos documentos
    documentos = set(cnm_df['Documento']) | set(soa_df['documento']) | set(sgp_df['CPF/CNPJ'])
    dados = []

    for doc in documentos:
        row_cnm = cnm_df[cnm_df['Documento'] == doc]
        row_soa = soa_df[soa_df['documento'] == doc]
        row_sgp = sgp_df[sgp_df['CPF/CNPJ'] == doc]

        presente_cnm = not row_cnm.empty
        presente_soa = not row_soa.empty
        presente_sgp = not row_sgp.empty

        id_val = row_soa['Unique ID'].values[0] if presente_soa else (int(row_cnm['Id'].values[0]) if presente_cnm else "-")
        nome = row_soa['devedor'].values[0] if presente_soa else "-"
        data = row_soa['Data Inclusão'].values[0] if presente_soa else (
                pd.to_datetime(row_cnm['Data / Hora'].values[0], dayfirst=True).strftime('%d/%m/%Y %H:%M') if presente_cnm else "-")

        local = f"<span style='color:{"green" if presente_cnm else "red"}'>CNM</span> | " +                 f"<span style='color:{"green" if presente_soa else "red"}'>SOA</span> | " +                 f"<span style='color:{"green" if presente_sgp else "red"}'>SGP</span>"

        # Lógica combinada
        if presente_sgp and (presente_cnm or presente_soa):
            tipo = "INCLUSÃO"
            status = "NEGATIVADO"
        elif presente_cnm and not presente_sgp:
            tipo = "EXCLUSAO"
            status = "BAIXADO"
        else:
            tipo = "ERRO"
            status = "ERRO"

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
    df.to_excel(os.path.join(saida_dir, "resultado_unificado.xlsx"), index=False)
    gerar_html(df)
    return df

def gerar_html(df):
    html_tabela = df.to_html(index=False, escape=False, table_id='tabela', classes='display')
    subtitulos = {
        "NEGATIVADO": "Clientes negativados no CNM ou SOA e presentes no SGP.",
        "BAIXADO": "Clientes excluídos no CNM ou no SOA e ausentes no SGP.",
        "ERRO": "Clientes com inconsistência de presença entre os sistemas CNM, SOA e SGP."
    }
    html_code = f"""<!DOCTYPE html>
<html lang='pt-BR'>
<head>
    <meta charset='UTF-8'><title>Dashboard Unificado</title>
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
    <h2>Dashboard Unificado - CNM + SOA x SGP</h2>
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
    with open(os.path.join(saida_dir, "dashboard_unificado.html"), "w", encoding="utf-8") as f:
        f.write(html_code)

if __name__ == '__main__':
    comparar_tudo()
