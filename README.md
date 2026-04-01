# Transparência SIGCON — v2

Automação do pipeline de transparência de convênios e instrumentos do SIGCON-MG. O projeto baixa automaticamente o Controle SEI do OneDrive corporativo, filtra os dados do SIGCON, faz download dos PDFs no SEI, envia os novos arquivos para o Google Drive e gera uma planilha consolidada com links para os inteiros teores.

---

## O que o projeto faz

O pipeline é executado em quatro etapas sequenciais:

**Etapa 0 — Download do Controle SEI (OneDrive)** (`baixar_controle_onedrive.py`)

- Autentica no Microsoft 365 via MSAL (PublicClientApplication — login interativo na 1ª execução, cache nas seguintes)
- Baixa o arquivo `Controle SEI.xlsx` do OneDrive corporativo via Microsoft Graph API
- Salva em `data/input/`

**Etapa 1 — Filtro do SIGCON**

- Lê `Consultas SIGCON.xlsx` exportada manualmente do QlikView e colocada em `data/input/`
- Filtra: remove situação `BLOQUEADO` e mantém apenas instrumentos com `Data.Publicação` dentro dos últimos 3 anos (inclusive ano atual)
- Usa os `Código.SIAFI` resultantes para filtrar o Controle SEI, obtendo apenas os instrumentos relevantes

**Etapa 2 — Download dos PDFs do SEI** (`baixar_arquivos_sei.py`)

- Acessa o SEI via Selenium com as credenciais do `.env`
- Baixa somente os PDFs dos instrumentos que passaram pelo filtro SIGCON
- Salva em `data/output/PDFs/`, pulando arquivos que já existem

**Etapa 3 — Upload para o Google Drive**

- Autentica no Google Drive via OAuth2
- Envia para a pasta `transparencia` apenas os PDFs locais que ainda não estão no Drive

**Etapa 4 — Cruzamento e exportação** (`monitoramento_instrumentos_sigcon.py`)

- Lista todos os PDFs da pasta `transparencia` no Drive
- Cruza: SIGCON filtrado ↔ Controle SEI (chave: `Código.SIAFI` / `Nº SIAFI_(SIGCON)`)
- Cruza: resultado ↔ Drive (chave: `Doc_autorizativo` = nome do PDF sem extensão)
- Constrói automaticamente a coluna `Inteiro.Teor.do.Instrumento.-.TransfereGov`:
  ```
  https://discricionarias.transferegov.sistema.gov.br/.../ResultadoDaConsultaDePropostaDetalharProposta.do?idProposta=<Código.União>&Usr=guest&Pwd=guest
  ```
- Preenche `Inteiro.teor.do.Instrumento.-.Sigcon` com o link do Drive quando estiver vazio
- Exporta planilha `.xlsx` formatada com as colunas na ordem especificada

---

## Colunas da planilha final

| Coluna | Origem |
|--------|--------|
| Unidade.Orçamentária … Fim.Vigência.Inicial | Consultas SIGCON.xlsx |
| Inteiro.Teor.do.Instrumento.-.TransfereGov | Construída via fórmula + `Código.União` |
| Inteiro.teor.do.Instrumento.-.Sigcon | SIGCON ou link do Drive (quando vazio) |
| Doc_autorizativo | Coluna `Instrumento` do Controle SEI |
| nome_pdf | Nome do arquivo PDF no Drive |
| id.y | ID do arquivo no Google Drive |
| drive_resource | URL de visualização no Google Drive |

---

## Pré-requisitos

- Python 3.10+
- Google Chrome + ChromeDriver compatível
- Conta Microsoft com acesso ao OneDrive corporativo (`cecad365`)
- Conta Google com acesso à pasta `transparencia` no Drive
- Credenciais OAuth2 do Google Cloud Console (`json.json`)
- `Consultas SIGCON.xlsx` exportada manualmente do QlikView em `data/input/`

### Instalação

```bash
pip install -r requirements.txt
```

---

## Configuração

1. Crie o `.env` a partir do modelo:

```bash
cp .env.example .env
```

2. Preencha o `.env`:

```env
# Proxy corporativo
HTTP_PROXY=http://USUARIO:SENHA@host-proxy:porta
HTTPS_PROXY=http://USUARIO:SENHA@host-proxy:porta

# Credenciais SEI MG
LOGIN_SEI=seu_cpf_ou_login
SENHA_SEI=sua_senha

# Caminho do ChromeDriver
CHROMEDRIVER_PATH=C:\caminho\para\chromedriver.exe

# Google Drive OAuth2
GOOGLE_CREDENTIALS_JSON=C:\caminho\para\json.json
GOOGLE_AUTH_EMAIL=seu_email@gmail.com
```

3. Coloque o arquivo `Consultas SIGCON.xlsx` (exportado do QlikView) em `data/input/`.

4. **Primeira execução:**
   - **Microsoft (OneDrive):** abre janela do navegador para login Microsoft 365. O token é cacheado em `data/.token_cache.bin`.
   - **Google Drive:** abre janela de autorização OAuth2. O token é salvo em `token.json`.

5. Verifique as constantes de colunas em `src/config.py` caso os cabeçalhos das planilhas sejam diferentes.

---

## Estrutura do projeto

```
├── main.py                                    # Ponto de entrada — pipeline completo
├── src/
│   ├── config.py                              # Constantes e configurações
│   ├── baixar_controle_onedrive.py            # Etapa 0: download Controle SEI via MSAL
│   ├── baixar_arquivos_sei.py                 # Etapa 2: download PDFs via Selenium
│   └── monitoramento_instrumentos_sigcon.py   # Etapas 3-4: upload, cruzamento, exportação
├── data/
│   ├── input/
│   │   ├── Controle SEI.xlsx                  # Baixado automaticamente do OneDrive
│   │   └── Consultas SIGCON.xlsx              # Exportar manualmente do QlikView
│   └── output/
│       └── PDFs/                              # PDFs baixados do SEI
├── archive/
│   └── sigcon_transparencia_11-11.py          # Versão anterior (referência)
├── docs/
├── .env                                       # Credenciais (não versionado)
├── .env.example                               # Modelo de configuração
├── json.json                                  # Credenciais OAuth2 Google (não versionado)
├── token.json                                 # Token Google Drive (não versionado)
└── requirements.txt
```

---

## Como executar

### Pipeline completo (recomendado)

```bash
python main.py
```

### Executar etapas individualmente

```bash
# Somente download do Controle SEI do OneDrive
python src/baixar_controle_onedrive.py

# Somente download dos PDFs do SEI (usa todos do Controle SEI, sem filtro)
python src/baixar_arquivos_sei.py

# Somente cruzamento, upload Drive e exportação
python src/monitoramento_instrumentos_sigcon.py
```

---

## Tecnologias

| Tecnologia | Uso |
|------------|-----|
| Selenium | Automação do navegador para acesso ao SEI |
| MSAL | Autenticação Microsoft para download do OneDrive via Graph API |
| Google Drive API (v3) | Upload e listagem de PDFs na pasta transparencia |
| Pandas | Leitura, filtro, cruzamento e manipulação dos dados |
| OpenPyXL | Exportação da planilha formatada |
| python-dotenv | Carregamento de variáveis de ambiente |
| requests | Download de PDFs do SEI e chamadas HTTP ao Graph API |

---

## Observações

- `Consultas SIGCON.xlsx` deve ser exportada manualmente do QlikView antes de rodar o pipeline. O nome do arquivo deve ser exatamente `Consultas SIGCON.xlsx`.
- `Controle SEI.xlsx` é baixado automaticamente do OneDrive a cada execução (sempre atualizado).
- Os PDFs já existentes em `data/output/PDFs/` são ignorados no download (idempotente).
- PDFs já existentes no Google Drive não são reenviados (idempotente).
- Projeto desenvolvido para uso na rede corporativa da SEPLAG-MG com proxy autenticado (Prodemge).
