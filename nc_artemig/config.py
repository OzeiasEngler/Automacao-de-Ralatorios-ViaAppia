"""Config pipeline NC Artemig (MG). Lote 50 — EAF CONSOL. regime=artemig."""

from __future__ import annotations

import os
from datetime import datetime
from pathlib import Path

_ARTEMIG_ROOT = Path(__file__).resolve().parent

# V = {DIR}\{stem PDF por apontamento}. Padrão macro Nas02.
DIR_BASE_FOTOS_KCOR_PADRAO = r"O:\NOTIFICAÇÕES DER\_LANÇAR\_02 - Arquivos Fotos"
DIR_BASE_FOTOS_KCOR = (os.environ.get("ARTEMIG_KCOR_DIR_FOTOS") or "").strip() or DIR_BASE_FOTOS_KCOR_PADRAO

LOTE_CONCESSIONARIA: dict[str, str] = {
    "50": "CONSOL",
}

RODOVIAS_POR_LOTE: dict[str, tuple[str, ...]] = {
    "50": ("MG 050", "BR 265", "BR 491"),
}

LOTES_MENU_ANALISE: list[tuple[str, str]] = [
    ("50", "Lote 50 — CONSOL (Artemig MG)"),
]

ASSETS_DIR = _ARTEMIG_ROOT / "assets"
MALHA_DIR = _ARTEMIG_ROOT / "malha"
TEMPLATES_DIR = _ARTEMIG_ROOT / "templates"
# Análise PDF (lote 50): template com col. A Tipo, B SH, V Nº CONSOL (PDF Artemig).
TEMPLATE_RELATORIO_ANALISE_PDF = ASSETS_DIR / "Template" / "Template_EAF_artemig.xlsx"

# Grupo EAF 50 = CONSOL (Artemig); trechos MG/BR — não usar mapa ARTESP (lote 13).
MAPA_EAF_POR_LOTE: dict[str, list] = {
    "50": [
        {
            "grupo": 50,
            "empresa": "CONSOL",
            # Extensão km do contrato Artemig CONSOL (PDF análise / atribuição trecho).
            "trechos": [
                {"rodovia": "MG 050", "km_ini": 57.6, "km_fim": 402.0},
                {"rodovia": "BR 265", "km_ini": 637.2, "km_fim": 659.5},
                {"rodovia": "BR 491", "km_ini": 0.0, "km_fim": 4.7},
            ],
        }
    ],
}
# Chave = nome EAF (igual MAPA_EAF.empresa); valor = responsável(is) técnico(s) permitido(s).
MAPA_RESPONSAVEL_TECNICO_POR_LOTE: dict[str, dict] = {
    "50": {"CONSOL": "Consol Engenharia"},
}

MODELO_KCOR_KRIA = ASSETS_DIR / "Template" / "templates" / "_Planilha Modelo Kcor-Kria_artemig.xlsx"
# Formato nome saída = macro Nas01: yyyyMMdd-hhmm - Exportar Kcor.xlsx
NOME_SAIDA_EXCEL_KCOR_PREFIXO = "Exportar Kcor"
NOME_SAIDA_EXCEL_KCOR_EXT = ".xlsx"

COL_KCOR_KRIA = {
    "NumItem": 1,           # A
    "Origem": 2,            # B — "0-QID" ou "Orgão Fiscalizador"
    "Motivo": 3,            # C — fixo "Conservação de Rotina"
    "Classificacao": 4,     # D — class_Kcor (ex.: Eng. QID)
    "Tipo": 5,              # E — kcor (patologia mapeada)
    "Rodovia": 6,           # F
    "KMi": 7,               # G
    "KMf": 8,               # H
    "Sentido": 9,           # I
    "Local": 10,            # J
    "Gestor": 11,           # K
    "Executores": 12,       # L
    "Data_Solicitacao": 13, # M
    "Data_Suspensao": 14,   # N
    "Dt_Inicio_Prog": 15,   # O
    "Dt_Fim_Prog": 16,      # P
    "Dt_Inicio_Exec": 17,   # Q
    "Dt_Fim_Exec": 18,      # R
    "Prazo": 19,            # S
    "Obs_Gestor": 20,       # T
    "Observacoes": 21,      # U
    "Diretorio": 22,        # V — pasta fotos
    "Arquivos": 23,         # W — arquivos por código NC (ex.: CE2607782.jpg)
    "Indicador": 24,        # X
    "Unidade": 25,          # Y
}


def get_rodovias_lote(lote: str) -> tuple[str, ...]:
    return RODOVIAS_POR_LOTE.get((lote or "").strip(), ())


def nome_saida_excel_kcor(dt: datetime | None = None) -> str:
    """Retorna nome do arquivo: yyyyMMdd-HHmm - Exportar Kcor.xlsx."""
    if dt is None:
        dt = datetime.now()
    return f"{dt:%Y%m%d-%H%M} - {NOME_SAIDA_EXCEL_KCOR_PREFIXO}{NOME_SAIDA_EXCEL_KCOR_EXT}"
