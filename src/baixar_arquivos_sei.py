import os
import time

import requests
import pandas as pd
from dotenv import load_dotenv
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

from config import (
    BASE_DIR, DIRETORIO_PDFS, PLANILHA_CONTROLE,
    URL_SEI, URL_BASE_PDF,
    TIMEOUT, SLEEP_LOGIN, SLEEP_PESQUISA, CHROME_OPTIONS,
    XPATH_TELA_LOGIN, XPATH_USUARIO, XPATH_SENHA, XPATH_ORGAO,
    XPATH_ORGAO_SCC, XPATH_BOTAO_LOGIN, XPATH_BODY,
    XPATH_PESQUISA, XPATH_IFRAME_DOC, XPATH_LINK_PDF,
    ENTER, ESC,
    COL_SEI_NUM_SEI, COL_SEI_INSTRUMENTO,
)

load_dotenv(BASE_DIR / ".env")


# ── Funções de automação ──────────────────────────────────────────────────────

def fazer_login(navegador):
#Preenche credenciais e efetua login no SEI.
    WebDriverWait(navegador, TIMEOUT).until(
        EC.presence_of_element_located((By.XPATH, XPATH_TELA_LOGIN))
    )
    navegador.find_element(By.XPATH, XPATH_USUARIO).send_keys(os.getenv("LOGIN_SEI"))
    navegador.find_element(By.XPATH, XPATH_SENHA).send_keys(os.getenv("SENHA_SEI"))
    navegador.find_element(By.XPATH, XPATH_ORGAO).click()
    navegador.find_element(By.XPATH, XPATH_ORGAO_SCC).click()
    time.sleep(SLEEP_LOGIN)
    navegador.find_element(By.XPATH, XPATH_BOTAO_LOGIN).click()
    WebDriverWait(navegador, TIMEOUT).until_not(
        EC.presence_of_element_located((By.XPATH, XPATH_TELA_LOGIN))
    )
    print("✅ Login efetuado com sucesso!")


def fechar_modal(navegador):
    """Fecha modal de boas-vindas do SEI via tecla ESC."""
    try:
        body = WebDriverWait(navegador, TIMEOUT).until(
            EC.element_to_be_clickable((By.XPATH, XPATH_BODY))
        )
        body.send_keys(ESC)
        print("✅ Modal fechada com sucesso!")
    except Exception as e:
        print(f"⚠️ Sem modal ou erro ao fechar: {e}")


def baixar_pdf(navegador, instrumento, caminho_pdf, num_sei):
    """Pesquisa o instrumento no SEI e salva o PDF em caminho_pdf."""
    navegador.switch_to.default_content()

    campo = WebDriverWait(navegador, TIMEOUT).until(
        EC.presence_of_element_located((By.XPATH, XPATH_PESQUISA))
    )
    campo.send_keys(str(instrumento))
    campo.send_keys(ENTER)
    time.sleep(SLEEP_PESQUISA)

    try:
        iframe = WebDriverWait(navegador, TIMEOUT).until(
            EC.presence_of_element_located((By.XPATH, XPATH_IFRAME_DOC))
        )
        navegador.switch_to.frame(iframe)
    except Exception:
        print(f"❌ Erro ao acessar iframe do doc — {num_sei}")
        return

    try:
        link_el = WebDriverWait(navegador, TIMEOUT).until(
            EC.presence_of_element_located((By.XPATH, XPATH_LINK_PDF))
        )
        href    = link_el.get_attribute("href")
        pdf_url = URL_BASE_PDF + href.split("controlador.php?")[1]

        cookies  = {c["name"]: c["value"] for c in navegador.get_cookies()}
        response = requests.get(pdf_url, cookies=cookies)

        with open(caminho_pdf, "wb") as f:
            f.write(response.content)

        print(f"✅ PDF {num_sei} baixado.")
    except Exception as e:
        print(f"❌ Erro no download de {instrumento}: {e}")


# ── Função principal ──────────────────────────────────────────────────────────

def executar(instrumentos_df=None):
    """
    Faz login no SEI e baixa os PDFs dos instrumentos.

    Parâmetros
    ----------
    instrumentos_df : pd.DataFrame | None
        DataFrame com colunas COL_SEI_NUM_SEI e COL_SEI_INSTRUMENTO.
        Se None, lê todos os registros de PLANILHA_CONTROLE.
    """
    os.makedirs(DIRETORIO_PDFS, exist_ok=True)

    if instrumentos_df is None:
        controle_sei   = pd.read_excel(PLANILHA_CONTROLE)
        instrumentos_df = controle_sei[
            [COL_SEI_NUM_SEI, COL_SEI_INSTRUMENTO]
        ].dropna(subset=[COL_SEI_INSTRUMENTO])

    instrumentos_df = instrumentos_df.dropna(subset=[COL_SEI_INSTRUMENTO])
    print(f"ℹ️  {len(instrumentos_df)} instrumento(s) na lista de download.")

    options = Options()
    for arg in CHROME_OPTIONS:
        options.add_argument(arg)

    navegador = webdriver.Chrome(
        service=Service(os.getenv("CHROMEDRIVER_PATH")),
        options=options,
    )

    try:
        navegador.get(URL_SEI)
        fazer_login(navegador)
        fechar_modal(navegador)

        for _, row in instrumentos_df.iterrows():
            num_sei     = row[COL_SEI_NUM_SEI]
            instrumento = row[COL_SEI_INSTRUMENTO]
            caminho_pdf = os.path.join(DIRETORIO_PDFS, f"{instrumento}.pdf")

            try:
                if os.path.exists(caminho_pdf):
                    print(f"⚠️  PDF {num_sei} já existe, pulando.")
                    continue
                baixar_pdf(navegador, instrumento, caminho_pdf, num_sei)
            except Exception as e:
                print(f"❌ Erro ao processar {num_sei} ({instrumento}): {e}")

    finally:
        navegador.quit()
        print("✅ Download de PDFs concluído.")


# ── Execução standalone ───────────────────────────────────────────────────────
if __name__ == "__main__":
    executar()
