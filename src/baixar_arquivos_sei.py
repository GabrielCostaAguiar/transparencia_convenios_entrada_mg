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
    BASE_DIR, DIRETORIO_PDFS, PLANILHA_LISTA_DOWNLOAD,
    URL_SEI, URL_BASE_PDF,
    TIMEOUT, SLEEP_LOGIN, SLEEP_PESQUISA, CHROME_OPTIONS,
    XPATH_TELA_LOGIN, XPATH_USUARIO, XPATH_SENHA, XPATH_ORGAO,
    XPATH_ORGAO_SCC, XPATH_BOTAO_LOGIN, XPATH_BODY,
    XPATH_PESQUISA, XPATH_IFRAME_DOC, XPATH_LINK_PDF,
    ENTER, ESC,
    COLUNAS_TA, COL_DOC_AUTORIZATIVO, COL_OBSERVACAO,
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

def _registrar_observacao(lista_df, idx, numero):
    """Acumula 'Arquivo X não encontrado' na coluna Observação da linha idx."""
    msg = f"Arquivo {numero} não encontrado"
    atual = lista_df.at[idx, COL_OBSERVACAO]
    lista_df.at[idx, COL_OBSERVACAO] = f"{atual}; {msg}" if atual else msg


def executar(lista_df=None):
    """
    Faz login no SEI e, para cada linha da lista de download, baixa:
      1. O PDF do Instrumento  → {Doc_autorizativo}.pdf
      2. Cada Termo Aditivo    → {int(ta_num)}_{sufixo}.pdf  (ex: 78162090_1.pdf)

    Quando um arquivo não é encontrado no SEI, registra na coluna Observação.
    Ao final, salva a lista atualizada em PLANILHA_LISTA_DOWNLOAD.

    Parâmetros
    ----------
    lista_df : pd.DataFrame | None
        DataFrame gerado por gerar_lista_download().
        Se None, lê de PLANILHA_LISTA_DOWNLOAD.
    """
    os.makedirs(DIRETORIO_PDFS, exist_ok=True)

    if lista_df is None:
        lista_df = pd.read_excel(PLANILHA_LISTA_DOWNLOAD)

    if COL_OBSERVACAO not in lista_df.columns:
        lista_df[COL_OBSERVACAO] = ""
    else:
        lista_df[COL_OBSERVACAO] = lista_df[COL_OBSERVACAO].fillna("")

    ta_cols_presentes = [c for c in COLUNAS_TA if c in lista_df.columns]

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

        for idx, row in lista_df.iterrows():
            instrumento = row.get(COL_DOC_AUTORIZATIVO)

            # ── Instrumento ───────────────────────────────────────────────────
            if pd.notna(instrumento):
                caminho_pdf = os.path.join(DIRETORIO_PDFS, f"{instrumento}.pdf")
                try:
                    if os.path.exists(caminho_pdf):
                        print(f"⚠️  Instrumento {instrumento} já existe, pulando.")
                    else:
                        baixar_pdf(navegador, instrumento, caminho_pdf, instrumento)
                        if not os.path.exists(caminho_pdf) or os.path.getsize(caminho_pdf) == 0:
                            _registrar_observacao(lista_df, idx, instrumento)
                except Exception:
                    _registrar_observacao(lista_df, idx, instrumento)
                    print(f"❌ Instrumento {instrumento} não encontrado no SEI.")

            # ── Termos Aditivos ───────────────────────────────────────────────
            for ta_col in ta_cols_presentes:
                ta_val = row.get(ta_col)
                if pd.isna(ta_val):
                    continue
                sufixo      = ta_col.split()[-1]
                ta_num      = int(float(ta_val))
                nome_pdf    = f"{ta_num}_{sufixo}.pdf"
                caminho_pdf = os.path.join(DIRETORIO_PDFS, nome_pdf)
                try:
                    if os.path.exists(caminho_pdf):
                        print(f"⚠️  {nome_pdf} já existe, pulando.")
                    else:
                        baixar_pdf(navegador, ta_num, caminho_pdf, f"{ta_col} ({ta_num})")
                        if not os.path.exists(caminho_pdf) or os.path.getsize(caminho_pdf) == 0:
                            _registrar_observacao(lista_df, idx, ta_num)
                except Exception:
                    _registrar_observacao(lista_df, idx, ta_num)
                    print(f"❌ {ta_col} ({ta_num}) não encontrado no SEI.")

    finally:
        navegador.quit()
        lista_df.to_excel(PLANILHA_LISTA_DOWNLOAD, index=False)
        print(f"✅ Download concluído. Lista atualizada em {PLANILHA_LISTA_DOWNLOAD.name}")


# ── Execução standalone ───────────────────────────────────────────────────────
if __name__ == "__main__":
    executar()
