import sys
import pandas as pd
from pathlib import Path

SRC = Path(__file__).parent / "src"
sys.path.insert(0, str(SRC))

from baixar_controle_onedrive import baixar_controle_sei_onedrive
from baixar_arquivos_sei import executar as baixar_pdfs_sei
from monitoramento_instrumentos_sigcon import (
    autenticar_drive,
    fazer_upload_pdfs,
    listar_pdfs_drive,
    carregar_bases,
    filtrar_sigcon,
    cruzar_dados,
    preencher_link_inteiro_teor,
    exportar_planilha,
    _id_pasta_drive,
)
from config import (
    COL_SIGCON_SIAFI,
    COL_SEI_SIAFI,
    COL_SEI_NUM_SEI,
    COL_SEI_INSTRUMENTO,
)


def section(titulo):
    print("\n" + "=" * 60)
    print(titulo)
    print("=" * 60)


# ── Etapa 0: Baixar Controle SEI do OneDrive ──────────────────────────────────
section("ETAPA 0 — Baixando Controle SEI do OneDrive")
baixar_controle_sei_onedrive()

# ── Etapa 1: Carregar planilhas e filtrar SIGCON ───────────────────────────────
section("ETAPA 1 — Carregando e filtrando dados")
controle_sei, consultas_sigcon = carregar_bases()
sigcon_filtrado = filtrar_sigcon(consultas_sigcon)

# Filtra Controle SEI apenas pelos SIAFIs presentes no SIGCON filtrado
siafi_filtrados = set(sigcon_filtrado[COL_SIGCON_SIAFI].astype(str).tolist())
controle_sei[COL_SEI_SIAFI] = controle_sei[COL_SEI_SIAFI].astype(str)
controle_filtrado = controle_sei[controle_sei[COL_SEI_SIAFI].isin(siafi_filtrados)]
instrumentos_para_baixar = (
    controle_filtrado[[COL_SEI_NUM_SEI, COL_SEI_INSTRUMENTO]]
    .dropna(subset=[COL_SEI_INSTRUMENTO])
    .reset_index(drop=True)
)
print(f"ℹ️  {len(instrumentos_para_baixar)} instrumento(s) correspondem ao filtro SIGCON.")

# ── Etapa 2: Download dos PDFs do SEI ────────────────────────────────────────
section("ETAPA 2 — Baixando PDFs do SEI")
baixar_pdfs_sei(instrumentos_para_baixar)

# ── Etapa 3: Upload dos novos PDFs para o Google Drive ────────────────────────
section("ETAPA 3 — Google Drive: upload dos PDFs")
service  = autenticar_drive()
pasta_id = _id_pasta_drive(service)
fazer_upload_pdfs(service)

# ── Etapa 4: Listar Drive + cruzar dados + exportar ──────────────────────────
section("ETAPA 4 — Cruzando dados e exportando planilha")
lista_instrumentos = listar_pdfs_drive(service)
df = cruzar_dados(sigcon_filtrado, controle_sei, lista_instrumentos)
df = preencher_link_inteiro_teor(df, service, pasta_id)
exportar_planilha(df)

print("\n" + "=" * 60)
print("✅ Pipeline finalizado com sucesso.")
print("=" * 60)
