from pathlib import Path
from datetime import date
from dotenv import load_dotenv

# ── Diretórios ────────────────────────────────────────────────────────────────
BASE_DIR       = Path(__file__).resolve().parent.parent

load_dotenv(BASE_DIR / ".env")
DOWNLOAD_DIR   = BASE_DIR / "data" / "input"
DIRETORIO_PDFS = BASE_DIR / "data" / "output" / "PDFs"
CAMINHO_TOKEN_CACHE = BASE_DIR / "data" / ".token_cache.bin"

# ── Planilhas de entrada ──────────────────────────────────────────────────────
PLANILHA_CONTROLE         = DOWNLOAD_DIR / "Controle SEI.xlsx"
PLANILHA_CONSULTAS_SIGCON = DOWNLOAD_DIR / "Consultas SIGCON.xls"

# ── Gmail API (OAuth2) ─────────────────────────────────────────────────────
CAMINHO_GMAIL_CREDENTIALS = BASE_DIR / "json.json"
CAMINHO_GMAIL_TOKEN       = BASE_DIR / "data" / ".gmail_token.json"

# ── Autenticação Microsoft (MSAL PublicClientApplication) ─────────────────────
# App ID público do Azure CLI — funciona em qualquer tenant sem registro próprio
AZURE_CLIENT_ID = "d3590ed6-52b3-4102-aeff-aad2292ab01c"

# Link de compartilhamento do Controle SEI.xlsx no OneDrive
CONTROLE_SEI_SHAREPOINT_URL = "https://cecad365-my.sharepoint.com/:x:/g/personal/m16900633_ca_mg_gov_br/IQDavaoigLEtS4nQs_JCschkAQiUl1Fh8hZN7KJTul3aQOI?e=DeIO1a"

# ── Filtros ───────────────────────────────────────────────────────────────────
ANO_ATUAL          = date.today().year
ANO_LIMITE         = ANO_ATUAL - 3
DATA_FILTRO        = list(range(ANO_LIMITE, ANO_ATUAL + 1))
SITUACAO_BLOQUEADO = "BLOQUEADO"

# ── Planilha de saída ─────────────────────────────────────────────────────────
PLANILHA_OUTPUT_PREFIX = f"Consultas SIGCON - Instrumentos de {ANO_LIMITE} até {ANO_ATUAL} - ATUALIZADO"
ABA_EXCEL              = "Consulta_SIGCON"
CAMINHO_FINAL = BASE_DIR / "data" / "output"

# ── URLs ──────────────────────────────────────────────────────────────────────
URL_SEI              = "https://www.sei.mg.gov.br/sip/login.php?sigla_orgao_sistema=GOVMG&sigla_sistema=SEI&infra_url=L3NlaS8="
URL_BASE_PDF         = "https://www.sei.mg.gov.br/sei/controlador.php?"
URL_TRANSFEREGOV_BASE = (
    "https://discricionarias.transferegov.sistema.gov.br/voluntarias/ConsultarProposta/"
    "ResultadoDaConsultaDePropostaDetalharProposta.do?idProposta="
)

# ── Google Drive ──────────────────────────────────────────────────────────────
PASTA_DRIVE  = "transparencia"
SCOPES_DRIVE = ["https://www.googleapis.com/auth/drive"]
TOKEN_PATH   = BASE_DIR / "token.json"

# ── Timeouts e esperas (segundos) ─────────────────────────────────────────────
TIMEOUT        = 10
SLEEP_LOGIN    = 1
SLEEP_PESQUISA = 3

# ── Selenium ──────────────────────────────────────────────────────────────────
CHROME_OPTIONS = ["--start-maximized"]

# ── XPaths — Login ────────────────────────────────────────────────────────────
XPATH_TELA_LOGIN  = '//*[@id="area-cards-login"]'
XPATH_USUARIO     = '//*[@id="txtUsuario"]'
XPATH_SENHA       = '//*[@id="pwdSenha"]'
XPATH_ORGAO       = '//*[@id="selOrgao"]'
XPATH_ORGAO_SCC   = '//*[@id="selOrgao"]/option[53]'
XPATH_BOTAO_LOGIN = '//*[@id="Acessar"]'
XPATH_BODY        = '/html/body'

# ── XPaths — Navegação / Download ─────────────────────────────────────────────
XPATH_PESQUISA   = '//*[@id="txtPesquisaRapida"]'
XPATH_IFRAME_DOC = '//*[@id="ifrVisualizacao"]'
XPATH_LINK_PDF   = '//*[@id="divArvoreInformacao"]/a'

# ── Teclas ────────────────────────────────────────────────────────────────────
ENTER = '\ue007'
ESC   = '\ue00c'

# ── Nomes de colunas — Controle SEI.xlsx ──────────────────────────────────────
COL_SEI_NUM_SEI     = "Nº_SEI"              # número do processo SEI (ex: 1234567-89.2024)
COL_SEI_SIAFI       = "Nº SIAFI_(SIGCON)"   # código SIAFI (chave de join com SIGCON)
COL_SEI_INSTRUMENTO = "Instrumento"          # número do instrumento (nome do PDF)

# ── Nomes de colunas — Consultas SIGCON.xlsx ──────────────────────────────────
COL_SIGCON_SIAFI             = "Código SIAFI"
COL_SIGCON_CODIGO_UNIAO      = "Código União"
COL_SIGCON_SITUACAO          = "Situação"
COL_SIGCON_DATA_PUB          = "Data Publicação"
COL_SIGCON_INTEIRO_SIGCON    = "Inteiro teor do Instrumento - Sigcon"
COL_SIGCON_INTEIRO_TRANSFERE = "Inteiro teor do Instrumento - TransfereGov"

# ── Nomes de colunas internas (join / drive) ──────────────────────────────────
COL_DOC_AUTORIZATIVO = "Doc_autorizativo"
COL_NOME_PDF         = "nome_pdf"
COL_ID_PDF           = "id.y"
COL_DRIVE_RESOURCE   = "drive_resource"

# ── Ordem das colunas na planilha final ───────────────────────────────────────
# Os nomes das colunas do SIGCON (QlikView) usam espaço como separador.
COLUNAS_FINAIS = [
    "Unidade Orçamentária",
    "Código SIGCON",
    "Código União",
    "Código Plano de Trabalho",
    "Código SIAFI",
    "SEI",
    "Instrumento",
    "Título",
    "Proponente",
    "Concedente",
    "Esfera Concedente",
    "Objeto",
    "Situação",
    "Tipo de Contrapartida",
    "Data Assinatura",
    "Data Publicação",
    "Início Vigência",
    "Término Vigência",
    "Valor Proponente",
    "Valor Concedente",
    "Valor Total Convênio",
    "Fim Vigência Inicial",
    COL_SIGCON_INTEIRO_TRANSFERE,   # "Inteiro teor do Instrumento - TransfereGov"
    COL_SIGCON_INTEIRO_SIGCON,      # "Inteiro teor do Instrumento - Sigcon"
    COL_DOC_AUTORIZATIVO,           # "Doc_autorizativo"
    COL_NOME_PDF,                   # "nome_pdf"
    COL_ID_PDF,                     # "id.y"
    COL_DRIVE_RESOURCE,             # "drive_resource"
]

# ── Estilo da planilha exportada ──────────────────────────────────────────────
EXCEL_FONTE       = "Tahoma"
EXCEL_TAMANHO     = 8
EXCEL_LARGURA_COL = 15
