import base64
from datetime import date
import requests
import msal

from config import (
    AZURE_CLIENT_ID,
    CONTROLE_SEI_SHAREPOINT_URL,
    PLANILHA_CONTROLE,
    CAMINHO_TOKEN_CACHE,
)


def baixar_controle_sei_onedrive():
    """Baixa o Controle SEI.xlsx do OneDrive corporativo via MSAL + Microsoft Graph API.
    Pula o download se o arquivo já foi baixado hoje (mesmo dia).
    """
    if PLANILHA_CONTROLE.exists():
        mod_date = date.fromtimestamp(PLANILHA_CONTROLE.stat().st_mtime)
        if mod_date == date.today():
            print("⚠️  Controle SEI.xlsx já foi baixado hoje, pulando.")
            return

    authority = "https://login.microsoftonline.com/common"
    scopes    = ["https://graph.microsoft.com/Files.Read.All"]

    # Carrega cache de token para evitar nova autenticação interativa
    cache = msal.SerializableTokenCache()
    if CAMINHO_TOKEN_CACHE.exists():
        cache.deserialize(CAMINHO_TOKEN_CACHE.read_text())

    app = msal.PublicClientApplication(
        AZURE_CLIENT_ID,
        authority=authority,
        token_cache=cache,
    )

    # Tenta token em cache primeiro
    accounts = app.get_accounts()
    result   = None
    if accounts:
        result = app.acquire_token_silent(scopes, account=accounts[0])

    # Se não há token válido, usa device code flow (sem redirect URI)
    if not result:
        flow  = app.initiate_device_flow(scopes=scopes)
        if "user_code" not in flow:
            raise RuntimeError(f"Falha ao iniciar device flow: {flow}")
        print("\n" + flow["message"])  # instrução: acesse aka.ms/devicelogin e insira o código
        result = app.acquire_token_by_device_flow(flow)

    # Persiste o cache atualizado
    if cache.has_state_changed:
        CAMINHO_TOKEN_CACHE.parent.mkdir(parents=True, exist_ok=True)
        CAMINHO_TOKEN_CACHE.write_text(cache.serialize())

    if "access_token" not in result:
        raise RuntimeError(
            f"Falha na autenticação Microsoft: {result.get('error_description', result)}"
        )

    # Converte o link de compartilhamento para um share-id do Graph API
    # Referência: https://learn.microsoft.com/pt-br/graph/api/shares-get
    b64 = base64.b64encode(CONTROLE_SEI_SHAREPOINT_URL.encode()).decode()
    b64 = b64.rstrip("=").replace("/", "_").replace("+", "-")
    share_id = "u!" + b64

    graph_url = f"https://graph.microsoft.com/v1.0/shares/{share_id}/driveItem/content"
    headers   = {"Authorization": f"Bearer {result['access_token']}"}

    resp = requests.get(graph_url, headers=headers, allow_redirects=True)
    resp.raise_for_status()

    PLANILHA_CONTROLE.parent.mkdir(parents=True, exist_ok=True)
    PLANILHA_CONTROLE.write_bytes(resp.content)

    print(f"✅ Controle SEI.xlsx baixado do OneDrive ({len(resp.content) // 1024} KB).")


if __name__ == "__main__":
    baixar_controle_sei_onedrive()
