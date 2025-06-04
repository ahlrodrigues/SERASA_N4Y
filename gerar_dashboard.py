import pandas as pd
import os
import zipfile

# Função para validar arquivo Excel (.xlsx válido como ZIP)
def validar_arquivo_excel(caminho):
    if not os.path.exists(caminho):
        raise FileNotFoundError(f"[ERRO] Arquivo não encontrado: {caminho}")
    if os.path.getsize(caminho) < 10240:
        raise ValueError(f"[ERRO] Arquivo muito pequeno (provavelmente inválido): {caminho}")
    try:
        with zipfile.ZipFile(caminho, 'r') as zf:
            if zf.testzip() is not None:
                raise zipfile.BadZipFile("Conteúdo ZIP inválido")
    except zipfile.BadZipFile as e:
        raise zipfile.BadZipFile(f"[ERRO] Arquivo .xlsx inválido ou corrompido: {e}")

# Caminhos dos arquivos
caminho_cnm = './download/Relatorio_CNM.xlsx'
caminho_sgp = './download/Relatorio_SGP.xlsx'
saida_excel = './output/Dashboard_CNM_SGP.xlsx'
saida_html = './output/dashboard_interativo.html'

# Validação dos arquivos
validar_arquivo_excel(caminho_cnm)
validar_arquivo_excel(caminho_sgp)

# Cria pasta de saída
os.makedirs('./output', exist_ok=True)

# Leitura dos arquivos
print('[INFO] Lendo arquivos...')
cnm_df = pd.read_excel(caminho_cnm, engine="openpyxl")
sgp_df = pd.read_excel(caminho_sgp, skiprows=8)

# Normalização dos campos
cnm_df['Documento'] = cnm_df['Documento'].astype(str).str.replace(r'\D', '', regex=True)
sgp_df['CPF/CNPJ'] = sgp_df['CPF/CNPJ'].astype(str).str.replace(r'\D', '', regex=True)

# Junção e status
print('[INFO] Comparando registros...')
merged_df = cnm_df.merge(sgp_df, left_on='Documento', right_on='CPF/CNPJ', how='inner')
merged_df['Status'] = merged_df['Tipo'].map({'INCLUSAO': 'NEGATIVADO', 'EXCLUSAO': 'BAIXADO'})
merged_df['Local'] = 'CNM | SGP'

# Dashboard final
dashboard_df = merged_df[['Id', 'Documento', 'Nome/Razão Social', 'Data / Hora', 'Operação', 'Tipo', 'Local', 'Status']].copy()
dashboard_df.rename(columns={'Nome/Razão Social': 'Nome'}, inplace=True)
dashboard_df = dashboard_df.head(1_000_000)

# Identificar exclusivos
cnm_exclusivo = cnm_df[~cnm_df['Documento'].isin(sgp_df['CPF/CNPJ'])].copy()
cnm_exclusivo['Local'] = 'CNM'

sgp_exclusivo = sgp_df[~sgp_df['CPF/CNPJ'].isin(cnm_df['Documento'])].copy()
sgp_exclusivo['Local'] = 'SGP'

# Exporta para Excel
print('[INFO] Salvando Excel...')
with pd.ExcelWriter(saida_excel) as writer:
    dashboard_df.to_excel(writer, sheet_name='Dashboard', index=False)
    cnm_exclusivo.to_excel(writer, sheet_name='Exclusivos_CNM', index=False)
    sgp_exclusivo.to_excel(writer, sheet_name='Exclusivos_SGP', index=False)

# Exporta para HTML interativo
print('[INFO] Salvando HTML interativo...')
html_tabela = dashboard_df.to_html(index=False, classes='display', table_id='dashboard', border=0)

html_template = f"""<!DOCTYPE html>
<html lang="pt-BR">
<head>
    <meta charset="UTF-8">
    <title>Dashboard CNM x SGP</title>
    <link rel="stylesheet" type="text/css"
          href="https://cdn.datatables.net/1.13.4/css/jquery.dataTables.min.css">
    <script src="https://code.jquery.com/jquery-3.7.0.min.js"></script>
    <script src="https://cdn.datatables.net/1.13.4/js/jquery.dataTables.min.js"></script>
</head>
<body>
    <h2>Dashboard CNM x SGP</h2>
    {html_tabela}
    <script>
        $(document).ready(function() {{
            $('#dashboard').DataTable({{
                paging: true,
                searching: true,
                ordering: true,
                pageLength: 25,
                language: {{
                    url: '//cdn.datatables.net/plug-ins/1.13.4/i18n/pt-BR.json'
                }}
            }});
        }});
    </script>
</body>
</html>"""

with open(saida_html, 'w', encoding='utf-8') as f:
    f.write(html_template)

print('[SUCESSO] Dashboard Excel e HTML salvos com sucesso.')
