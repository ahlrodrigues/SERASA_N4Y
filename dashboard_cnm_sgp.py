import pandas as pd
import os

# Caminhos
caminho_dir = './download'
saida_dir = './output'
os.makedirs(saida_dir, exist_ok=True)

caminho_cnm = os.path.join(caminho_dir, 'Relatorio_CNM.xlsx')
caminho_sgp = os.path.join(caminho_dir, 'Relatorio_SGP.xlsx')

def normalizar_documento(coluna):
    return (
        coluna.astype(str)
            .str.replace(r'\D', '', regex=True)
            .str.strip()
    )

def definir_status(row, documentos_sgp):
    doc = row['Documento']
    tipo = str(row['Tipo']).strip().upper()
    if tipo == "INCLUSAO":
        return "NEGATIVADO"
    elif tipo == "EXCLUSAO":
        return "ERRO" if doc in documentos_sgp else "BAIXADO"
    else:
        return "ERRO"

def determinar_status(presente_sgp, tipo):
    tipo = tipo.strip().upper()
    if tipo.startswith("PF") or tipo.startswith("PJ"):
        return "SCORE"
    elif tipo == "INCLUSAO":
        return "NEGATIVADO" if presente_sgp else "ERRO"
    elif tipo == "EXCLUSAO":
        return "ERRO" if presente_sgp else "BAIXADO"
    else:
        return "ERRO"

def carregar_dados():
    if not os.path.exists(caminho_cnm) or not os.path.exists(caminho_sgp):
        print('[ERRO] Arquivo CNM ou SGP não encontrado.')
        exit(1)

    print('[INFO] Lendo arquivos CNM e SGP...')
    cnm_df = pd.read_excel(caminho_cnm, engine='openpyxl')
    sgp_df = pd.read_excel(caminho_sgp, skiprows=8)

    print('[INFO] Normalizando documentos...')
    cnm_df['Documento'] = normalizar_documento(cnm_df['Documento'])
    sgp_df['CPF/CNPJ'] = normalizar_documento(sgp_df['CPF/CNPJ'])

    documentos_sgp = set(sgp_df['CPF/CNPJ'])

    print('[INFO] Aplicando status ao CNM...')
    cnm_df['Status'] = cnm_df.apply(lambda row: definir_status(row, documentos_sgp), axis=1)

    return cnm_df, sgp_df

def gerar_dashboard():
    cnm_df, sgp_df = carregar_dados()
    documentos = set(cnm_df['Documento']) | set(sgp_df['CPF/CNPJ'])

    dados = []
    for doc in documentos:
        row_cnm = cnm_df[cnm_df['Documento'] == doc]
        row_sgp = sgp_df[sgp_df['CPF/CNPJ'] == doc]
        presente_sgp = not row_sgp.empty

        if not row_cnm.empty:
            tipo = row_cnm['Tipo'].values[0]
            data = pd.to_datetime(row_cnm['Data / Hora'].values[0], dayfirst=True).strftime('%d/%m/%Y %H:%M')
        else:
            tipo = "-"
            data = "-"

        nome = row_cnm['Nome'].values[0] if not row_cnm.empty and 'Nome' in row_cnm else "-"
        id_val = "-"
        local = " | ".join([
    f"<span style='color:{'green' if not row.empty else 'red'}'>{origem}</span>"
    for origem, row in [('CNM', row_cnm), ('SGP', row_sgp)]
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
    df.to_excel(os.path.join(saida_dir, "resultado_cnm_sgp.xlsx"), index=False)
    gerar_html(df)

def gerar_html(df):
    html_tabela = df.to_html(index=False, escape=False, table_id='tabela', classes='display')
    subtitulos = {
        "NEGATIVADO": "Clientes negativados no CNM e presentes no SGP.",
        "BAIXADO": "Clientes excluídos no CNM e ausentes no SGP.",
        "ERRO": "Clientes com inconsistência entre CNM e SGP."
    }
    html_code = f"""<!DOCTYPE html>
<html lang='pt-BR'>
<head>
    <meta charset='UTF-8'><title>Dashboard CNM x SGP</title>
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
    <h2>Dashboard CNM x SGP</h2>
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
                    const colunasParaFiltrar = {{
                        3: "Data",
                        4: "Tipo",
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
            table.on('search.dt', function () {{
                const val = table.column(6, {{search: 'applied'}}).data()[0];
                $('#subtitulo').html(subtitulos[val] || '');
            }});
        }});
    </script>
</body>
</html>"""
    with open(os.path.join(saida_dir, "dashboard_cnm_sgp.html"), "w", encoding="utf-8") as f:
        f.write(html_code)

if __name__ == '__main__':
    gerar_dashboard()
