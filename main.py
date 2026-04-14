import sys
from pathlib import Path

SRC = Path(__file__).parent / "src"
sys.path.insert(0, str(SRC))

from baixar_controle_onedrive import baixar_controle_sei_onedrive
from baixar_arquivos_sei import executar as baixar_pdfs_sei
from monitoramento_instrumentos_sigcon import (
    autenticar_drive,
    fazer_upload_pdfs,
    listar_pdfs_drive,
    listar_pdfs_ta_drive,
    carregar_bases,
    filtrar_sigcon,
    gerar_lista_download,
    cruzar_dados,
    preencher_link_inteiro_teor,
    exportar_planilha,
    _id_pasta_drive,
)


def section(titulo):
    print("\n" + "=" * 60)
    print(titulo)
    print("=" * 60)


# ── Etapa 0: Baixar Controle SEI do OneDrive ──────────────────────────────────
section("ETAPA 0 — Baixando Controle SEI do OneDrive")
baixar_controle_sei_onedrive()

# ── Etapa 1: Carregar planilhas, filtrar SIGCON e gerar lista de download ──────
section("ETAPA 1 — Carregando, filtrando e gerando lista de download")
controle_sei, consultas_sigcon = carregar_bases()
sigcon_filtrado = filtrar_sigcon(consultas_sigcon)
lista_df = gerar_lista_download(sigcon_filtrado, controle_sei)

# ── Etapa 2: Download dos PDFs do SEI (Instrumentos + Termos Aditivos) ────────
section("ETAPA 2 — Baixando PDFs do SEI")
baixar_pdfs_sei(lista_df)

# ── Etapa 3: Upload dos novos PDFs para o Google Drive ────────────────────────
section("ETAPA 3 — Google Drive: upload dos PDFs")
service  = autenticar_drive()
pasta_id = _id_pasta_drive(service)
fazer_upload_pdfs(service)

# ── Etapa 4: Listar Drive + cruzar dados + exportar ──────────────────────────
section("ETAPA 4 — Cruzando dados e exportando planilha")
lista_instrumentos = listar_pdfs_drive(service)
lista_ta_drive     = listar_pdfs_ta_drive(service)
df = cruzar_dados(sigcon_filtrado, controle_sei, lista_instrumentos, lista_ta_drive)
df = preencher_link_inteiro_teor(df, service, pasta_id)
exportar_planilha(df)

print("\n" + "=" * 60)
print("✅ Pipeline finalizado com sucesso.")
print("=" * 60)
