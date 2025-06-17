# 📊 Dashboard SERASA - Automação, Comparação e Visualização de Relatórios

Este projeto automatiza o download, análise e visualização de relatórios de negativação (CNM, SGP e SOA), gerando dashboards interativos em HTML. Ele utiliza Python, Selenium, Pandas e Jinja2 para facilitar o controle e monitoramento das movimentações de documentos.

---

## ⚙️ Funcionalidades

- Acesso automatizado a portais SOA, CNM e SGP
- Download automático de arquivos `.csv` e `.xlsx`
- Processamento e normalização dos dados
- Geração de dashboards interativos com filtros, busca e ordenação
- Comparações entre relatórios (ex: status NEGATIVADO, BAIXADO, ERRO)
- Exportação em Excel e HTML

---

## 📁 Estrutura de Pastas

```bash
/download         # Arquivos baixados automaticamente pelos scripts
/output           # Dashboards HTML e arquivos Excel gerados
.env              # Arquivo com credenciais (não subir para o GitHub)
gerar_dashboard.py
baixar_dados_selenium.py
```

---

## 📦 Requisitos

### 1. Python (>= 3.10)

Instale as dependências com:

```bash
pip install -r requirements.txt
```

Ou manualmente:

```bash
pip install pandas openpyxl xlrd selenium python-dotenv jinja2
```

### 2. ChromeDriver

- Baixe o [ChromeDriver](https://sites.google.com/a/chromium.org/chromedriver/) compatível com sua versão do Google Chrome.
- Coloque o executável em `./chromedriver` ou no PATH do sistema.

### 3. Arquivo `.env`

Crie um arquivo `.env` na raiz do projeto com:

```dotenv
USUARIO_SOA=seu_usuario_soa
SENHA_SOA=sua_senha_soa
USUARIO_CNM=seu_usuario_cnm
SENHA_CNM=sua_senha_cnm
USUARIO_SGP=seu_usuario_sgp
SENHA_SGP=sua_senha_sgp
```

---

## 📥 Arquivos Esperados

Os scripts esperam encontrar ou baixar os seguintes arquivos:

- `/download/Relatorio_CNM.xlsx`
- `/download/Relatorio_SGP.xlsx`
- `/download/Ativas.csv`
- `/download/Baixadas.csv`
- `/download/Pendentes.csv`
- `/download/Determinacao.csv`
- `/download/Erros.csv`

> Os arquivos são baixados automaticamente ao rodar o script Selenium, e renomeados conforme necessário.

---

## ▶️ Como usar

1. Configure o `.env`
2. Execute o script de automação (opcional, para baixar os arquivos):

```bash
python baixar_dados_selenium.py
```

3. Gere o dashboard com:

```bash
python gerar_dashboard.py
```

4. O resultado será salvo em `/output/dashboard_interativo.html`.

---

## 🛡️ Segurança

Este projeto **não deve ter o `.env` com credenciais enviado para repositórios públicos**. Use `.gitignore` para garantir isso:

```
.env
```

---

## 📄 Licença

Este projeto é livre para uso interno e educacional. Para uso comercial, consulte os termos de uso das plataformas automatizadas (SOA, CNM, SGP).

---

## 🙋‍♂️ Suporte

Em caso de dúvidas, sugestões ou melhorias, sinta-se à vontade para abrir uma **issue** ou um **pull request**.