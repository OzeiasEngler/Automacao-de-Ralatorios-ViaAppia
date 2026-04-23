#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
GERADOR DE RELATÓRIO ARTESP - CORE ENGINE
==========================================
Versão: 3.8.3 (com correções de normalização)

Correções aplicadas:
  - normalizar_unidade() aplicada nas properties do GeoJSON
  - processar_local() normaliza espaços → underscores
  - "um" mapeado para "un" no MAPA_UNIDADES
  - Locais com espaço (MARGINAL NORTE) → underscore (MARGINAL_NORTE)
"""

# ══════════════════════════════════════
#  IMPORTS
# ══════════════════════════════════════
import os
import sys
import json
import math
import shutil
import datetime
import tempfile
import hashlib
import hmac
import base64
import platform
import warnings
import unicodedata
import re
import html
import calendar
from collections import Counter, defaultdict, OrderedDict
from decimal import Decimal, ROUND_HALF_UP
import uuid

warnings.filterwarnings("ignore")

try:
    import pandas as pd
    PANDAS_OK = True
except ImportError:
    PANDAS_OK = False

try:
    from dateutil.relativedelta import relativedelta
    DATEUTIL_OK = True
except ImportError:
    DATEUTIL_OK = False

try:
    import openpyxl
    from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
    from openpyxl.utils import get_column_letter
    OPENPYXL_OK = True
except ImportError:
    OPENPYXL_OK = False

try:
    import jsonschema
    JSONSCHEMA_OK = True
except ImportError:
    JSONSCHEMA_OK = False

try:
    from reportlab.lib.pagesizes import A4, landscape
    from reportlab.lib import colors
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.platypus import (
        SimpleDocTemplate, Paragraph, Spacer,
        Table, TableStyle, PageBreak
    )
    from reportlab.lib.units import cm, mm
    REPORTLAB_OK = True
except ImportError:
    REPORTLAB_OK = False

try:
    import tkinter as tk
    from tkinter import filedialog, messagebox, simpledialog
    TK_OK = True
except ImportError:
    TK_OK = False
    messagebox = None

# ══════════════════════════════════════
#  VERSÃO E CONSTANTES
# ══════════════════════════════════════
VERSAO = "3.8.3"

# Diretórios
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))

# Parâmetros de geometria
ESPACO_MINIMO_M = 5
TOL_SIMPLIF_M = 2
MAX_PONTOS_POR_LINHA = 5000
MAX_SALTO_M = 100  # L21: se dist > N m entre sequenciais, não ligue (novo segmento). Também evita outlier.
TOLERANCIA_QUEBRA_IMPOSSIVEL_M = 500  # se distância física > (delta_km*1000 + N), quebra linha (MultiLineString)
# Deslocamento geográfico (offset) para separar pistas Norte/Sul no mapa (evita encavalar no canteiro central)
# 0.0005 deg ≈ 55 m — separação visível; aumente (ex.: 0.001 ≈ 110 m) se ainda parecer colado
OFFSET_SENTIDO_GRAUS = 0.0005
MAX_ANGULO_ANTIBICO = 110  # graus: remove 'bicos' quando a mudança de direção é >= este valor (110–150 evita cortar curvas reais)
CORRIGIR_EIXO = False         # Ativar/desativar correção (desativado para não quebrar geração; reativar após validar _aplicar_correcao_eixo_ao_cache)
INTERVALO_CORRECAO = 15       # metros entre pontos
JANELA_MEDIANA = 50           # pontos para filtro
TOLERANCIA_LATERAL = 15       # metros de tolerância
JANELA_SUAVIZACAO = 5         # pontos para suavização


def corrigir_malha_dataframe(df, janela_mediana=None, janela_suavizacao=None, tolerancia=None):
    """
    Corrige zigzag/suaviza o eixo (anti-zigzag). Preserva Km.
    Agrupa por Rodovia (+ Sentido, + Segmento se existirem), aplica mediana móvel e suavização.
    """
    if df is None or df.empty:
        return df
    janela_mediana = janela_mediana if janela_mediana is not None else JANELA_MEDIANA
    janela_suavizacao = janela_suavizacao if janela_suavizacao is not None else JANELA_SUAVIZACAO
    tolerancia = tolerancia if tolerancia is not None else TOLERANCIA_LATERAL

    df = df.copy()
    df["Km"] = pd.to_numeric(df["Km"], errors="coerce")
    df["Latitude"] = pd.to_numeric(df["Latitude"], errors="coerce")
    df["Longitude"] = pd.to_numeric(df["Longitude"], errors="coerce")
    df = df.dropna(subset=["Rodovia", "Km", "Latitude", "Longitude"])

    tem_sentido = "Sentido" in df.columns
    tem_segmento = "Segmento" in df.columns
    chaves = ["Rodovia"]
    if tem_sentido:
        chaves.append("Sentido")
    if tem_segmento:
        chaves.append("Segmento")

    out = []
    for _, g in df.groupby(chaves, dropna=False):
        g = g.sort_values("Km").reset_index(drop=True)
        lat_med = g["Latitude"].rolling(window=janela_mediana, center=True, min_periods=1).median()
        lon_med = g["Longitude"].rolling(window=janela_mediana, center=True, min_periods=1).median()
        if tolerancia is not None and tolerancia > 0:
            dlat = (g["Latitude"] - lat_med).abs()
            dlon = (g["Longitude"] - lon_med).abs()
            mask_ok = (dlat <= tolerancia) & (dlon <= tolerancia)
            if mask_ok.mean() > 0.7:
                g = g[mask_ok].reset_index(drop=True)
                lat_med = g["Latitude"].rolling(window=janela_mediana, center=True, min_periods=1).median()
                lon_med = g["Longitude"].rolling(window=janela_mediana, center=True, min_periods=1).median()
        lat_s = lat_med.rolling(window=janela_suavizacao, center=True, min_periods=1).mean()
        lon_s = lon_med.rolling(window=janela_suavizacao, center=True, min_periods=1).mean()
        g = g.copy()
        g["Latitude"] = lat_s
        g["Longitude"] = lon_s
        out.append(g)

    df2 = pd.concat(out, ignore_index=True)
    sort_cols = ["Rodovia"]
    if tem_sentido:
        sort_cols.append("Sentido")
    if tem_segmento:
        sort_cols.append("Segmento")
    sort_cols.append("Km")
    df2 = df2.sort_values(sort_cols).reset_index(drop=True)
    return df2


def _aplicar_correcao_eixo_ao_cache():
    """Aplica correção de eixo (mediana + suavização) aos pontos já carregados no CACHE.
    Trabalha em cópia e só atualiza o cache no final, para não corromper em caso de erro."""
    if not CORRIGIR_EIXO or not CACHE.dados:
        return
    novo_dados = dict(CACHE.dados)
    for key in list(novo_dados.keys()):
        pts = novo_dados[key]
        if not pts or len(pts) < 2:
            continue
        parts = key.split("|", 2)
        rod = (parts[0].strip() if len(parts) > 0 else "") or ""
        sentido = (parts[1].strip() if len(parts) > 1 else "") or None
        local = (parts[2].strip() if len(parts) > 2 else "") or None
        rows = [
            {
                "Rodovia": rod,
                "Sentido": sentido,
                "Km": p["km"],
                "Latitude": p["lat"],
                "Longitude": p["lon"],
            }
            for p in pts
        ]
        df = pd.DataFrame(rows)
        df_corr = corrigir_malha_dataframe(
            df,
            janela_mediana=JANELA_MEDIANA,
            janela_suavizacao=JANELA_SUAVIZACAO,
            tolerancia=TOLERANCIA_LATERAL,
        )
        if df_corr is None or df_corr.empty:
            continue
        df_corr = df_corr.dropna(subset=["Km", "Latitude", "Longitude"])
        if df_corr.empty or len(df_corr) < 2:
            continue
        novos_pts = [
            {"km": float(row["Km"]), "lon": float(row["Longitude"]), "lat": float(row["Latitude"])}
            for _, row in df_corr.iterrows()
        ]
        if novos_pts:
            novo_dados[key] = novos_pts
    CACHE.dados.clear()
    CACHE.dados.update(novo_dados)
    print("  [OK] Correção de eixo aplicada (mediana + suavização).")


# Parâmetros do Excel
CABECALHO_LINHAS = 5
LINHA_INICIO_DADOS = 6

# Linhas de exemplo do template (para remoção automática)
LINHAS_EXEMPLO_TEMPLATE = []

# Unidades válidas do schema
UNIDADES_VALIDAS = {"km", "vb", "un", "t", "kg", "m", "m2", "m3", "H/h", "L", "k"}

# ══════════════════════════════════════════════════════
#  MAPA DE UNIDADES — com TODAS as variações possíveis
# ══════════════════════════════════════════════════════
MAPA_UNIDADES = {
    "un": "un", "und": "un", "unid": "un", "unidade": "un", "uni": "un",
    "um": "un",                                               # ← CORREÇÃO: "um" → "un"
    "u": "un", "pç": "un", "pc": "un", "peca": "un", "peça": "un",
    "m": "m", "mt": "m", "metro": "m", "metros": "m", "mts": "m",
    "m2": "m2", "m²": "m2", "m 2": "m2", "metro2": "m2", "metros2": "m2",
    "m3": "m3", "m³": "m3", "m 3": "m3", "metro3": "m3", "metros3": "m3",
    "km": "km", "quilometro": "km", "quilometros": "km",
    "l": "L", "lt": "L", "lts": "L", "litro": "L", "litros": "L",  # ← CORREÇÃO: minúsculo → "L"
    "vb": "vb", "verba": "vb", "gb": "vb", "global": "vb",
    "t": "t", "ton": "t", "tonelada": "t", "toneladas": "t",
    "kg": "kg", "quilo": "kg", "quilos": "kg", "quilograma": "kg",
    "hh": "H/h", "h/h": "H/h", "homemhora": "H/h", "hm": "H/h", "k": "k",
}

# Locais válidos do schema
LOCAIS_VALIDOS = [
    "PISTA_NORTE", "PISTA_SUL", "PISTA_LESTE", "PISTA_OESTE",
    "CANTEIRO_CENTRAL", "EIXO_LESTE-OESTE", "EIXO_NORTE-SUL",
    "PISTA_EXTERNA", "PISTA_INTERNA",
    "MARGINAL_SUL", "MARGINAL_NORTE", "MARGINAL_LESTE", "MARGINAL_OESTE",
    "DISPOSITIVO", "ALÇA", "INTERLIGAÇÃO",
    "FX_DOM", "FX_NAO_EDIF",
]

# Mapeamento de local para sentido na malha (geometria)
# Malhas padronizadas: A=Rodovia, B=Km, C=Sentido (Crescente/Decrescente), D=Latitude, E=Longitude
# Lote pode ter Norte, Sul, Leste, Oeste (e Externa/Interna onde aplicável)
MAPA_LOCAL_PARA_SENTIDO_MALHA = {
    "PISTA_NORTE": "Norte", "PISTA_SUL": "Sul", "PISTA_LESTE": "Leste", "PISTA_OESTE": "Oeste",
    "MARGINAL_NORTE": "Norte", "MARGINAL_SUL": "Sul", "MARGINAL_LESTE": "Leste", "MARGINAL_OESTE": "Oeste",
    "PISTA_EXTERNA": "Externa", "PISTA_INTERNA": "Interna",
    "CANTEIRO_CENTRAL": "Norte", "DISPOSITIVO": "Norte",
    "EIXO_NORTE-SUL": "Norte", "EIXO_LESTE-OESTE": "Leste",
    "ALÇA": "Norte", "INTERLIGAÇÃO": "Norte", "FX_DOM": "Norte", "FX_NAO_EDIF": "Norte",
    "CRESCENTE": "Crescente", "DECRESCENTE": "Decrescente",
    "Crescente": "Crescente", "Decrescente": "Decrescente",
}
# Sentido malha -> local (para gerar uma feature por sentido)
SENTIDO_MALHA_PARA_LOCAL = {
    "Norte": "PISTA_NORTE", "Sul": "PISTA_SUL", "Leste": "PISTA_LESTE", "Oeste": "PISTA_OESTE",
    "Externa": "PISTA_EXTERNA", "Interna": "PISTA_INTERNA",
}
# Mapeamento para extrair_sentido (gerar_id) — igual ao script mãe
MAPA_SENTIDO = {
    "PISTA_NORTE": "N", "PISTA_SUL": "S", "PISTA_LESTE": "L", "PISTA_OESTE": "O",
    "MARGINAL_NORTE": "N", "MARGINAL_SUL": "S", "MARGINAL_LESTE": "L", "MARGINAL_OESTE": "O",
    "EIXO_NORTE-SUL": "N/S", "EIXO_LESTE-OESTE": "L/O",
    "PISTA_INTERNA": "I", "PISTA_EXTERNA": "E",
    "CANTEIRO_CENTRAL": "CC", "DISPOSITIVO": "D", "ALÇA": "A",
    "INTERLIGAÇÃO": "INT", "FX_DOM": "FX", "FX_NAO_EDIF": "FX",
}

MAPA_SENTIDO = {
    "N": "Norte", "S": "Sul", "L": "Leste", "O": "Oeste",
    "NORTE": "Norte", "SUL": "Sul", "LESTE": "Leste", "OESTE": "Oeste",
    "CRESCENTE": "Crescente", "DECRESCENTE": "Decrescente",
    "C": "Crescente", "D": "Decrescente",
}

# Cores dos marcadores (alfinetes) — igual ao script mãe
CORES_MARCADORES = {
    "inicial": "#1976D2",
    "final": "#388E3C",
    "ponto": "#D32F2F",
}

# ══════════════════════════════════════
#  MODALIDADES E SCHEMAS
# ══════════════════════════════════════
MODALIDADES = {
    "1": {"chave": "conservacao", "rotulo": "Conservacao", "schema_asset": ("schema", "conserva.schema.r0.json")},
    "2": {"chave": "obras", "rotulo": "Obras", "schema_asset": ("schema", "obras.schema.r0.json")},
}

# ══════════════════════════════════════
#  VERSÕES DE RELATÓRIO
# ══════════════════════════════════════
VERSOES_RELATORIO = OrderedDict([
    ("r0", {"rotulo": "R0 — Programação Anual", "tipo": "ANUAL", "senha": False}),
    ("r1", {"rotulo": "R1 — Revisão 1", "tipo": "ANUAL", "senha": False}),
    ("r2", {"rotulo": "R2 — Revisão 2", "tipo": "ANUAL", "senha": False}),
    ("r3", {"rotulo": "R3 — Revisão 3", "tipo": "ANUAL", "senha": False}),
    ("m", {"rotulo": "M — Programação Mensal", "tipo": "MENSAL", "senha": False}),
    ("e", {"rotulo": "E — Executado", "tipo": "EXECUTADO", "senha": False}),
    ("c1", {"rotulo": "C1 — Correção 1", "tipo": "ANUAL", "senha": True}),
    ("c2", {"rotulo": "C2 — Correção 2", "tipo": "ANUAL", "senha": True}),
])

SENHA_CORRECAO = "artesp2024"

# Meses
MESES_ABREV = {
    1: "jan", 2: "fev", 3: "mar", 4: "abr", 5: "mai", 6: "jun",
    7: "jul", 8: "ago", 9: "set", 10: "out", 11: "nov", 12: "dez",
}

MESES_NOME_COMPLETO = {
    1: "Janeiro", 2: "Fevereiro", 3: "Marco", 4: "Abril",
    5: "Maio", 6: "Junho", 7: "Julho", 8: "Agosto",
    9: "Setembro", 10: "Outubro", 11: "Novembro", 12: "Dezembro",
}

# Nomes dos meses para o nome do arquivo (minúsculo, com acento em março)
MESES_NOME_SAIDA = {
    1: "janeiro", 2: "fevereiro", 3: "março", 4: "abril",
    5: "maio", 6: "junho", 7: "julho", 8: "agosto",
    9: "setembro", 10: "outubro", 11: "novembro", 12: "dezembro",
}

# Lotes válidos no schema
LOTES_SCHEMA_VALIDOS = [
    "L01", "L06", "L07", "L11", "L13", "L16", "L19", "L20", "L21", "L22",
    "L23", "L24", "L25", "L27", "L28", "L29", "L30", "L31", "L32", "L33", "L34"
]

# Rodovias por lote
RODOVIAS_POR_LOTE = {
    "L13": [
        "SP0000075", "SP0000127", "SP0000280", "SP0000300", "SPI102300",
    ],
    "L21": [
        "AHB000146", "BRE000005", "BRE000232", "BTC000040", "BTC000055", "BTC000260", "BTC000353",
        "CHS000326", "CHS000387", "CPR000010", "CPR000152", "ESF000020", "HRT000050", "HTR000050", "IDT000085",
        "ITN000313", "LEP000030", "LEP000119", "LEP000148", "LEP000321", "LEP000357", "LEP000363",
        "LEP000374", "LRP000321", "MOR000040", "MOR000137", "MOR000293", "MTB000070", "MTB000148",
        "MTB000195", "PFZ000080", "PIR000030", "PRD000010", "RFR000154", "RPD000015", "RPD000020",
        "SP0000101", "SP0000113", "SP0000209", "SP0000300", "SP0000308",
        "SPA007209", "SPA022101", "SPA026101", "SPA032101", "SPA043101", "SPA051101", "SPA139308",
        "SPA155308", "SPA159300", "SPA172300", "SPA176300", "SPA193300", "SPA196300", "SPA231300",
        "SPA241300", "SPA251300", "SPA270300", "SPA283300", "SPI162308", "SPI181300", "TIT000366",
    ],
    "L26": ["SP0000021"],
}

# Lotes que usam geometria por rodovia (sem sentido/local no eixo).
# L13 removido: malha tem coluna Local igual à L21 — cache por local funciona normalmente.
LOTE_GEOMETRIA_POR_RODOVIA_SO = {
    "L26": True,
}

# Semântica por lote: sentido de exibição → sentido na malha (Crescente/Decrescente).
# Suporta Norte/Sul, Leste/Oeste e Externa/Interna conforme malha de cada lote.
LOTE_SENTIDO_PARA_MALHA = {
    "L13": {
        "Norte": "Crescente", "Sul": "Decrescente",
        "Leste": "Crescente", "Oeste": "Decrescente",
        "Externa": "Crescente", "Interna": "Decrescente",
        "Crescente": "Crescente", "Decrescente": "Decrescente",
    },
    "L21": {
        "Norte": "Decrescente", "Sul": "Crescente",
        "Leste": "Crescente", "Oeste": "Decrescente",
        "Externa": "Crescente", "Interna": "Decrescente",
        "Crescente": "Crescente", "Decrescente": "Decrescente",
    },
    "L26": {
        "Norte": "Crescente", "Sul": "Decrescente",
        "Leste": "Crescente", "Oeste": "Decrescente",
        "Externa": "Crescente", "Interna": "Decrescente",
        "Crescente": "Crescente", "Decrescente": "Decrescente",
    },
}

# Sentido para exibição nas features por lote (Crescente/Decrescente da malha → rótulo na saída)
LOTE_SENTIDO_DISPLAY = {
    "L13": {"Crescente": "Norte", "Decrescente": "Sul"},
    "L21": {"Crescente": "Externa", "Decrescente": "Interna"},
    "L26": {"Crescente": "Norte", "Decrescente": "Sul"},
}

# Sentido na malha por rodovia (sobrescreve o do lote quando definido). Ex.: SP0000127 Norte = Decrescente.
# Formato: rodovia normalizada → { "Norte"|"Sul"|"Leste"|"Oeste": "Crescente"|"Decrescente" }
RODOVIA_SENTIDO_PARA_MALHA = {
    "SP0000021": {
        "Crescente": "Crescente", "Decrescente": "Decrescente",
        "Norte": "Crescente", "Sul": "Decrescente",
        "Leste": "Crescente", "Oeste": "Decrescente",
        "Externa": "Crescente", "Interna": "Decrescente",
    },
    "SP0000075": {"Norte": "Crescente", "Sul": "Decrescente", "Leste": "Crescente", "Oeste": "Decrescente"},
    "SP0000300": {"Leste": "Decrescente", "Oeste": "Crescente", "Norte": "Crescente", "Sul": "Decrescente"},
    "SP0000280": {"Oeste": "Crescente", "Leste": "Decrescente", "Norte": "Crescente", "Sul": "Decrescente"},
    "SP0000127": {"Norte": "Decrescente", "Sul": "Crescente", "Leste": "Crescente", "Oeste": "Decrescente"},
    "SPI102300": {"Leste": "Decrescente", "Oeste": "Crescente", "Norte": "Crescente", "Sul": "Decrescente"},
}

LOTE_USAR_ALFINETE_QUANDO_KM_IGUAL = True

# Lotes disponíveis (nomes exatos dos arquivos em assets/malha - Linux é case-sensitive)
LOTES = {
    "13": {"sigla": "L13", "rotulo": "Lote 13", "eixo": ("malha", "Eixo Lote 13.xlsx")},
    "21": {"sigla": "L21", "rotulo": "Lote 21", "eixo": ("malha", "Eixo lote 21.xlsx")},
    "26": {"sigla": "L26", "rotulo": "Lote 26", "eixo": ("malha", "Eixo Lote 26.csv")},
}

# Desenvolvedor
DESENVOLVEDOR_NOME = "Ozeias Engler"
DESENVOLVEDOR_EMAIL = "oseias.engler@hotmail.com"
DESENVOLVEDOR_TELEFONE = "(11) 98755-4794"
NOME_ARQUIVO_LIC = "licenca.lic"

# ══════════════════════════════════════
#  CACHE GLOBAL
# ══════════════════════════════════════
DADOS_MALHA = {}


class CacheMalha:
    def __init__(self):
        self.dados = {}
        self.tem_sentidos = False
        self._rodovias = set()
        self._rod_sentidos = {}
        self.sentidos_invertidos_detectados = []  # [(rod, sentido), ...] para log/relatório

    def limpar(self):
        self.dados.clear()
        self.tem_sentidos = False
        self._rodovias.clear()
        self._rod_sentidos.clear()
        self.sentidos_invertidos_detectados = []

    def finalizar_carregamento(self):
        """Verifica se a progressão de km de cada rod|sentido|local bate com o sentido declarado.
        Inversões são corrigidas automaticamente em obter_intervalo(); aqui só registra para log/auditoria.
        """
        self.sentidos_invertidos_detectados = []
        for key, pts in self.dados.items():
            if not pts or "|" not in key:
                continue
            parts = key.split("|", 2)
            rod = (parts[0].strip() if len(parts) > 0 else "") or ""
            sentido = (parts[1].strip() if len(parts) > 1 else "") or None
            if not sentido:
                continue
            kms = [p["km"] for p in pts]
            is_crescente_real = kms == sorted(kms)
            is_decrescente_real = kms == sorted(kms, reverse=True)
            sentido_low = str(sentido).strip().lower()
            if sentido_low in ("crescente", "cresc") and not is_crescente_real:
                self.sentidos_invertidos_detectados.append((rod, sentido))
                print(f"  [AVISO] Sentido invertido detectado (corrigido automaticamente): {rod} — {sentido}")
            elif sentido_low in ("decrescente", "decresc") and not is_decrescente_real:
                self.sentidos_invertidos_detectados.append((rod, sentido))
                print(f"  [AVISO] Sentido invertido detectado (corrigido automaticamente): {rod} — {sentido}")

        # Ordenar cada lista por km (Crescente → asc, Decrescente → desc) para evitar zig-zag
        for key in list(self.dados.keys()):
            if not key or "|" not in key:
                continue
            parts = key.split("|", 2)
            sentido = (parts[1].strip() if len(parts) > 1 else "") or None
            reverse = (
                sentido is not None
                and str(sentido).strip().lower() in ("decrescente", "decresc")
            )
            self.dados[key] = sorted(self.dados[key], key=lambda p: p["km"], reverse=reverse)

    def _chave(self, rod, sentido, local=None):
        """Chave inclui sentido para não misturar Norte/Sul na hora de buscar geometria."""
        return f"{rod}|{sentido or ''}|{local or ''}"

    def adicionar(self, rod, sentido, km, lon, lat, local=None):
        """Armazena ponto com chave rod|sentido|local (uma feature por par Sentido+Local)."""
        key = self._chave(rod, sentido, local)
        if key not in self.dados:
            self.dados[key] = []
        self.dados[key].append({"km": km, "lon": lon, "lat": lat})
        self._rodovias.add(rod)
        if rod not in self._rod_sentidos:
            self._rod_sentidos[rod] = set()
        if sentido:
            self._rod_sentidos[rod].add(sentido)
            self.tem_sentidos = True

    def contem(self, rod, sentido=None, local=None):
        if sentido is not None or local is not None:
            if local is not None and sentido is not None:
                return self._chave(rod, sentido, local) in self.dados
            prefix = f"{rod}|{sentido or ''}|{local or ''}"
            for k in self.dados:
                if k == prefix or (local is None and k.startswith(f"{rod}|{sentido}|")) or (sentido is None and k.startswith(f"{rod}|") and k.endswith(f"|{local or ''}")):
                    return True
            return False
        for k in self.dados:
            if k.startswith(f"{rod}|"):
                return True
        return rod in self._rodovias

    def sentidos_disponiveis(self, rod):
        return sorted(self._rod_sentidos.get(rod, set()))

    def pares_sentido_local_disponiveis(self, rod):
        """Retorna [(sentido, local), ...] para essa rodovia (agrupamento Sentido+Local)."""
        pares = set()
        prefix = f"{rod}|"
        for key in self.dados:
            if not key.startswith(prefix):
                continue
            parts = key.split("|", 2)
            sent = (parts[1] if len(parts) > 1 else "") or None
            loc = (parts[2] if len(parts) > 2 else "") or None
            pares.add((sent, loc))
        return sorted(pares, key=lambda x: (x[0] or "", x[1] or ""))

    def obter(self, rod, sentido=None, local=None):
        if sentido is not None and local is not None:
            return self.dados.get(self._chave(rod, sentido, local), [])
        if sentido is not None:
            pts = []
            prefix = f"{rod}|{sentido}|"
            for k, v in self.dados.items():
                if k.startswith(f"{rod}|") and (k == f"{rod}|{sentido}|" or k.startswith(prefix)):
                    pts.extend(v)
            return sorted(pts, key=lambda p: p["km"])
        pts = []
        for k, v in self.dados.items():
            if k.startswith(f"{rod}|"):
                pts.extend(v)
        return sorted(pts, key=lambda p: p["km"])

    def obter_intervalo(self, rod, km_ini, km_fim, sentido=None, local=None):
        """Pontos no intervalo [km_ini, km_fim] para rod+sentido+local.
        Ordem: Crescente → km crescente; Decrescente → km decrescente.
        """
        chave = self._chave(rod, sentido, local)
        pts = self.obter(rod, sentido, local)
        if not pts:
            return []
        ki, kf = min(km_ini, km_fim), max(km_ini, km_fim)
        filtrado = [p for p in pts if ki - 0.05 <= p["km"] <= kf + 0.05]
        if not filtrado and pts:
            # Fallback: pontos mais próximos
            pts_sorted = sorted(pts, key=lambda p: min(abs(p["km"] - ki), abs(p["km"] - kf)))
            filtrado = pts_sorted[:2]
        # Decrescente/Oeste → ordem km decrescente; Crescente/Leste → km crescente
        reverse = (
            sentido is not None
            and str(sentido).strip().lower() in ("decrescente", "decresc", "oeste")
        )
        return sorted(filtrado, key=lambda p: p["km"], reverse=reverse)

    def resumo_rodovias_km(self):
        """Retorna lista de dict com rodovia, km_min e km_max para diagnóstico (ex.: quando nenhuma feature é gerada)."""
        out = []
        for rod in sorted(self._rodovias):
            pts = self.obter(rod, None, None)
            if not pts:
                continue
            kms = [p["km"] for p in pts]
            out.append({"rodovia": rod, "km_min": min(kms), "km_max": max(kms)})
        return out


CACHE = CacheMalha()

# Locais que geram LineString (Pista/Marginal). DISPOSITIVO e ALÇA não seguem KM linear da rodovia:
# idealmente tratar como Point/MultiPoint ou features individuais, não LineString por KM.
LOCAIS_PRINCIPAIS_LINHA = ["PISTA_NORTE", "PISTA_SUL", "MARGINAL_NORTE", "MARGINAL_SUL"]


# ══════════════════════════════════════
#  FUNÇÕES DE DIRETÓRIO E ASSETS
# ══════════════════════════════════════
def _base_dir_app():
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))


def _inicializar_diretorios():
    base = _base_dir_app()
    for d in ["assets", os.path.join("assets", "schema"), os.path.join("assets", "template"),
              os.path.join("assets", "malha"), "saida"]:
        p = os.path.join(base, d)
        if not os.path.isdir(p):
            try:
                os.makedirs(p, exist_ok=True)
            except OSError:
                pass


def _bases_asset():
    if getattr(sys, "frozen", False):
        meipass = getattr(sys, "_MEIPASS", None)
        exe_dir = os.path.dirname(sys.executable)
        if meipass and os.path.isdir(meipass):
            yield meipass
        if exe_dir and os.path.isdir(exe_dir):
            yield exe_dir
        return
    base = _base_dir_app()
    yield base


def _path_asset(*parts):
    for base in _bases_asset():
        if os.path.basename(base).lower() == 'assets':
            c = os.path.join(base, *parts)
            if os.path.exists(c):
                return c
            c2 = os.path.join(os.path.dirname(base), 'assets', *parts)
            if os.path.exists(c2):
                return c2
            return c
        c = os.path.join(base, 'assets', *parts)
        if os.path.exists(c):
            return c
        c2 = os.path.join(base, *parts)
        if os.path.exists(c2):
            return c2
    return os.path.join(_base_dir_app(), 'assets', *parts)


def _path_asset_eixo(*parts):
    path = _path_asset(*parts)
    if os.path.exists(path):
        return path
    nome_arquivo = parts[-1] if parts else ""
    base_nome, _ = os.path.splitext(nome_arquivo)
    dir_parts = list(parts[:-1]) if len(parts) > 1 else list(parts)
    for base in _bases_asset():
        if "lote 21" in base_nome.lower() or "lote21" in base_nome.lower().replace(" ", ""):
            for d in [os.path.join(base, "assets", *dir_parts), os.path.join(base, *dir_parts)]:
                c = os.path.join(d, "Eixo lote 21 corrigido.csv")
                if os.path.exists(c):
                    return c
        if os.path.basename(base).lower() == "assets":
            dirs = [os.path.join(base, *parts[:-1]), os.path.join(os.path.dirname(base), "assets", *parts[:-1])]
        else:
            dirs = [os.path.join(base, "assets", *parts[:-1]), os.path.join(base, *parts[:-1])]
        for diretorio in dirs:
            for extensao in (".geojson", ".xlsx", ".csv"):
                candidato = os.path.join(diretorio, base_nome + extensao)
                if os.path.exists(candidato):
                    return candidato
    return path


# ══════════════════════════════════════
#  NORMALIZAÇÃO DE TEXTOS
# ══════════════════════════════════════
def _strip_accents(s):
    if not s:
        return ""
    return "".join(c for c in unicodedata.normalize("NFKD", str(s)) if not unicodedata.combining(c))


def norm_key(s):
    s = _strip_accents(s or "").strip().lower()
    return s.replace("\n", " ").replace("\r", " ").replace("_", "").replace("-", "").replace(" ", "")


def normalizar_lote(lote_str):
    if not lote_str or pd.isna(lote_str):
        return None
    s = str(lote_str).strip().upper().replace("LOTE", "").replace("LOT", "").replace(" ", "")
    if s.isdigit():
        return f"L{int(s)}"
    if s.startswith("L"):
        nums = "".join(filter(str.isdigit, s))
        if nums:
            return f"L{int(nums)}"
    return s


def normalizar_item(item):
    r"""Normaliza o campo item para cumprir o schema: ^[a-z](\.\d+)+$"""
    if item is None or (isinstance(item, float) and math.isnan(item)):
        return ""
    s = str(item).strip().lower()
    s = s.replace(" ", "")
    s = re.sub(r"\.[a-z]$", "", s)
    s = re.sub(r"\.+$", "", s)
    s = re.sub(r"[^a-z0-9\.]", "", s)
    s = re.sub(r"\.{2,}", ".", s)
    return s


def normalizar_rodovia(nome):
    if pd.isna(nome):
        return ""
    limpo = str(nome).upper().replace("-", "").replace(" ", "").replace("/", "").strip()
    # Normalização de segurança para Hortolândia (HRT → HTR)
    limpo = limpo.replace("HRT", "HTR")
    if "SPD" in limpo:
        m = re.match(r"(\d{1,2})SPD(\d+)", limpo)
        if m:
            return f"{int(m.group(1)):02d}SPD{int(m.group(2)):06d}"
    if limpo.startswith("SPM"):
        m = re.match(r"SPM(\d+)([A-Z])?", limpo)
        if m:
            return f"SPM{int(m.group(1)):05d}{m.group(2) or ''}"
    if limpo.startswith("SPI"):
        m = re.match(r"SPI(\d+)([A-Z])?", limpo)
        if m:
            return f"SPI{int(m.group(1)):06d}{m.group(2) or ''}"
    if limpo.startswith("SP") and not limpo.startswith(("SPA", "SPI", "SPM", "SPD")):
        nums = "".join(filter(str.isdigit, limpo))
        return f"SP{int(nums):07d}" if nums else limpo
    return limpo


_RODOVIAS_ARTESP_CODIGOS_SENT = object()
_RODOVIAS_ARTESP_CODIGOS = _RODOVIAS_ARTESP_CODIGOS_SENT


def _carregar_codigos_rodovias_xlsx_interno():
    if not PANDAS_OK:
        return None
    path = _path_asset("malha", "rodovias.xlsx")
    if not os.path.isfile(path):
        return None
    try:
        df = pd.read_excel(path, sheet_name=0, header=0)
    except Exception:
        return None
    col = None
    for c in df.columns:
        if str(c).strip().lower() == "codigo":
            col = c
            break
    if col is None:
        return None
    out = set()
    for v in df[col].dropna():
        n = normalizar_rodovia(v)
        if n:
            out.add(n)
    return frozenset(out) if out else None


def obter_codigos_rodovias_validos():
    """Códigos normalizados da coluna ``codigo`` em ``assets/malha/rodovias.xlsx`` (lista oficial). None se indisponível."""
    global _RODOVIAS_ARTESP_CODIGOS
    if _RODOVIAS_ARTESP_CODIGOS is _RODOVIAS_ARTESP_CODIGOS_SENT:
        _RODOVIAS_ARTESP_CODIGOS = _carregar_codigos_rodovias_xlsx_interno()
    return _RODOVIAS_ARTESP_CODIGOS


# ══════════════════════════════════════════════════════════
#  NORMALIZAR UNIDADE — ← CORREÇÃO PRINCIPAL
# ══════════════════════════════════════════════════════════
def normalizar_unidade(raw):
    """
    Normaliza unidade para valor aceito pelo schema.
    Converte variações como 'um'→'un', 'l'→'L', 'und'→'un', etc.
    """
    if raw is None or (isinstance(raw, float) and math.isnan(raw)):
        return None
    s = str(raw).strip()
    # Caso exato já válido
    if s in UNIDADES_VALIDAS:
        return s
    # Lookup no mapa (minúsculo)
    s_lower = s.lower().strip()
    s_lower = re.sub(r"\s+", "", s_lower)
    if s_lower in MAPA_UNIDADES:
        return MAPA_UNIDADES[s_lower]
    # Tentar sem acentos
    s_clean = _strip_accents(s_lower)
    if s_clean in MAPA_UNIDADES:
        return MAPA_UNIDADES[s_clean]
    # Tentar só alfanuméricos
    s_alpha = re.sub(r"[^a-z0-9/]", "", s_lower)
    if s_alpha in MAPA_UNIDADES:
        return MAPA_UNIDADES[s_alpha]
    # Retornar original (vai falhar na validação, mas não perde dado)
    return s


# ══════════════════════════════════════════════════════════════
#  PROCESSAR LOCAL — igual ao script mãe (greedy match em LOCAIS_VALIDOS)
# ══════════════════════════════════════════════════════════════
def processar_local(val):
    """
    Igual ao script mãe: usa greedy match em LOCAIS_VALIDOS.
    Aceita "MARGINAL LESTE; MARGINAL OESTE", "MARGINAL_LESTE_MARGINAL_OESTE", etc.
    """
    if val is None or (isinstance(val, float) and math.isnan(val)):
        return []
    s = ";".join([str(x).strip() for x in val]) if isinstance(val, list) else str(val)
    s = s.upper().strip()
    if not s:
        return []
    s = s.replace(" ", "_")  # normalizar espaço→underscore para match com LOCAIS_VALIDOS
    out = []
    for loc in sorted(LOCAIS_VALIDOS, key=len, reverse=True):
        while loc in s:
            if loc not in out:
                out.append(loc)
            s = s.replace(loc, " ", 1)
    return out if out else []


# ══════════════════════════════════════
#  FUNÇÕES DE COORDENADAS
# ══════════════════════════════════════
def _parse_coord_malha(v):
    if v is None or (isinstance(v, float) and math.isnan(v)):
        return None
    s = str(v).strip().replace(",", ".")
    if not s or s in ("", "nan"):
        return None
    if s.count(".") <= 1:
        try:
            return float(s)
        except (ValueError, TypeError):
            return None
    s_limpa = s.replace(".", "").replace("-", "")
    if not s_limpa.isdigit():
        return None
    try:
        num = int(s.replace(".", "").replace(" ", ""))
        return num / 1e8
    except (ValueError, TypeError):
        return None


def _distancia_haversine(p1, p2):
    """Distância em metros entre dois pontos [lon, lat]."""
    R = 6371000
    lon1, lat1 = math.radians(p1[0]), math.radians(p1[1])
    lon2, lat2 = math.radians(p2[0]), math.radians(p2[1])
    dlat = lat2 - lat1
    dlon = lon2 - lon1
    a = math.sin(dlat / 2) ** 2 + math.cos(lat1) * math.cos(lat2) * math.sin(dlon / 2) ** 2
    return R * 2 * math.atan2(math.sqrt(a), math.sqrt(1 - a))


# ══════════════════════════════════════
#  FUNÇÕES KM E FORMATAÇÃO
# ══════════════════════════════════════
def sha256_arquivo(p):
    h = hashlib.sha256()
    try:
        with open(p, "rb") as f:
            for ch in iter(lambda: f.read(1024 * 1024), b""):
                h.update(ch)
        return h.hexdigest()
    except (OSError, IOError):
        return "ERRO_HASH"


def _parse_km_excel(v):
    if v is None or (isinstance(v, float) and math.isnan(v)):
        return None
    s = str(v).strip().replace(",", ".")
    if "+" in s:
        try:
            parts = s.split("+", 1)
            km = float(parts[0].strip() or 0)
            m = float(parts[1].strip() or 0)
            return km + m / 1000.0
        except (ValueError, IndexError):
            return None
    try:
        return float(re.sub(r"[^\d\.\-]", "", s) or 0)
    except (ValueError, TypeError):
        return None


def _to_float(v):
    if v is None or (isinstance(v, float) and math.isnan(v)):
        return None
    if isinstance(v, (int, float)):
        return float(v)
    try:
        return float(re.sub(r"[^\d\.\-]", "", str(v).replace(",", ".")))
    except (ValueError, TypeError):
        return None


def _snap_km(val):
    if val is None:
        return None
    return float(Decimal(str(val)).quantize(Decimal("0.001"), rounding=ROUND_HALF_UP))


def _formatar_km_relatorio(val):
    if val is None or (isinstance(val, float) and (val != val)):
        return ""
    km = _snap_km(val)
    if km is None:
        return ""
    return f"{km:.3f}".replace(".", ",")


def _formatar_data_iso(v):
    """Converte data para formato ISO yyyy-mm-dd. Aceita datetime, dd/mm/yyyy, yyyy-mm-dd ou número serial (Excel)."""
    if v is None or (isinstance(v, float) and math.isnan(v)):
        return None
    if isinstance(v, (datetime.date, datetime.datetime)):
        return v.strftime("%Y-%m-%d")
    # Número serial Excel (ex.: 45321)
    try:
        if isinstance(v, (int, float)) and 1 <= v <= 2958465:
            from datetime import timedelta
            base = datetime.datetime(1899, 12, 30)
            d = base + timedelta(days=int(v))
            return d.strftime("%Y-%m-%d")
    except (ValueError, OverflowError, OSError):
        pass
    s = str(v).strip()
    if not s:
        return None
    # Tentar dd/mm/yyyy
    m = re.match(r"(\d{1,2})/(\d{1,2})/(\d{4})", s)
    if m:
        return f"{m.group(3)}-{int(m.group(2)):02d}-{int(m.group(1)):02d}"
    # Tentar yyyy-mm-dd
    m = re.match(r"(\d{4})-(\d{1,2})-(\d{1,2})", s)
    if m:
        return f"{m.group(1)}-{int(m.group(2)):02d}-{int(m.group(3)):02d}"
    return None


def _formatar_data_saida_dma(v):
    """Formata data para dd/mm/yyyy."""
    if v is None:
        return ""
    if isinstance(v, str):
        m = re.match(r"(\d{4})-(\d{2})-(\d{2})", v)
        if m:
            return f"{m.group(3)}/{m.group(2)}/{m.group(1)}"
    if isinstance(v, (datetime.date, datetime.datetime)):
        return v.strftime("%d/%m/%Y")
    return str(v)


def _to_string_required(v, default=""):
    if v is None or (isinstance(v, float) and math.isnan(v)):
        return default
    s = str(v).strip()
    return s if s else default


def _to_string_or_null(v):
    if v is None or (isinstance(v, float) and math.isnan(v)):
        return None
    s = str(v).strip()
    return s if s else None


# ══════════════════════════════════════
#  GEOMETRIA
# ══════════════════════════════════════
def _dedup(coords):
    out, last = [], None
    for c in coords:
        if c and len(c) >= 2:
            lon, lat = float(c[0]), float(c[1])
            if last is None or abs(lon - last[0]) > 1e-9 or abs(lat - last[1]) > 1e-9:
                out.append([lon, lat])
                last = (lon, lat)
    return out


def _filtrar_salto_geografico(coords, max_salto_m=MAX_SALTO_M):
    """Remove pontos que 'pulam' mais que max_salto_m do anterior (outlier que causa zig-zag)."""
    if not coords or len(coords) < 2 or max_salto_m <= 0:
        return coords
    res = [coords[0]]
    for p in coords[1:]:
        if _distancia_haversine(res[-1], p) <= max_salto_m:
            res.append(p)
    if res[-1] != coords[-1] and len(res) >= 2:
        res.append(coords[-1])  # manter último ponto da linha
    return res


def _calcular_azimute(p1, p2):
    """Ângulo entre dois pontos [lon, lat] em graus (aproximado para distâncias curtas)."""
    lon1, lat1 = float(p1[0]), float(p1[1])
    lon2, lat2 = float(p2[0]), float(p2[1])
    delta_x = lon2 - lon1
    delta_y = lat2 - lat1
    if abs(delta_x) < 1e-12 and abs(delta_y) < 1e-12:
        return 0.0
    return math.degrees(math.atan2(delta_x, delta_y))


def _filtrar_mudanca_brusca(coords, max_angulo=MAX_ANGULO_ANTIBICO):
    """Remove pontos que formam 'bico' (mudança de direção >= max_angulo graus). Evita triângulos pontiagudos de erro de GPS."""
    if not coords or len(coords) < 3:
        return coords
    res = [coords[0]]
    for i in range(1, len(coords) - 1):
        p_ant = res[-1]
        p_atual = coords[i]
        p_prox = coords[i + 1]
        angulo_entrada = _calcular_azimute(p_ant, p_atual)
        angulo_saida = _calcular_azimute(p_atual, p_prox)
        diff = abs(angulo_saida - angulo_entrada)
        if diff > 180:
            diff = 360 - diff
        if diff < max_angulo:
            res.append(p_atual)
    res.append(coords[-1])
    return res


def filtrar_espaco_minimo(coords, esp=ESPACO_MINIMO_M):
    if not coords or len(coords) < 2:
        return coords
    res = [coords[0]]
    ult = coords[0]
    for p in coords[1:-1]:
        if _distancia_haversine(ult, p) >= esp:
            res.append(p)
            ult = p
    if coords[-1] != res[-1]:
        res.append(coords[-1])
    return res


def _perp_dist_m(pt, s, e):
    lon, lat = pt
    lat_m = (s[1] + e[1]) / 2
    mlon = 111320 * math.cos(math.radians(lat_m))
    mlat = 111320
    x, y = (lon - s[0]) * mlon, (lat - s[1]) * mlat
    x2, y2 = (e[0] - s[0]) * mlon, (e[1] - s[1]) * mlat
    if x2 == 0 and y2 == 0:
        return math.hypot(x, y)
    t = max(0, min(1, (x * x2 + y * y2) / (x2 * x2 + y2 * y2)))
    return math.hypot(x - t * x2, y - t * y2)


def _rdp(coords, tol):
    n = len(coords)
    if n < 3:
        return coords
    keep = [False] * n
    keep[0] = keep[-1] = True
    stack = [(0, n - 1)]
    while stack:
        s, e = stack.pop()
        if e - s < 2:
            continue
        mx, mi = -1, -1
        for i in range(s + 1, e):
            d = _perp_dist_m(coords[i], coords[s], coords[e])
            if d > mx:
                mx, mi = d, i
        if mi >= 0 and mx > tol:
            keep[mi] = True
            stack.append((s, mi))
            stack.append((mi, e))
    return [coords[i] for i, k in enumerate(keep) if k]


def simplificar(coords, tol=TOL_SIMPLIF_M):
    if not coords or len(coords) < 3 or not tol:
        return coords
    return _rdp(coords, tol)


def _limitar_pontos_linha(coords, max_pts=MAX_PONTOS_POR_LINHA):
    if not coords or len(coords) <= max_pts:
        return coords
    n = len(coords)
    indices = [0] + [int(round(i * (n - 1) / (max_pts - 1))) for i in range(1, max_pts - 1)] + [n - 1]
    return [list(coords[i]) for i in indices]


def _normalizar_geoms(geoms):
    """Remove geometrias None e normaliza."""
    if not geoms:
        return []
    return [g for g in geoms if g is not None]


# ══════════════════════════════════════
#  SALVAR GEOJSON
# ══════════════════════════════════════
def _simplificar_coordenadas(coords, step):
    """Reduz pontos: mantém primeiro, último e cada step-ésimo ponto. step=1 = sem redução."""
    if step <= 1 or not coords:
        return coords
    # LineString: coords = [[lon, lat], [lon, lat], ...]
    if isinstance(coords[0], (list, tuple)) and len(coords[0]) >= 2 and isinstance(coords[0][0], (int, float)):
        n = len(coords)
        if n <= 2:
            return coords
        idx = list(range(0, n, step))
        if idx and idx[-1] != n - 1:
            idx.append(n - 1)
        return [coords[i] for i in idx]
    # MultiLineString: coords = [ [[lon,lat],...], [[lon,lat],...], ... ]
    if isinstance(coords[0], (list, tuple)) and isinstance(coords[0][0], (list, tuple)):
        return [_simplificar_coordenadas(linha, step) for linha in coords]
    return coords


def salvar_geojson(caminho, obj):
    # step=2 mantém cada 2º ponto (reduz tamanho). Use ARTESP_GEOJSON_SIMPLIFY_STEP=1 para manter todos os pontos.
    step = max(1, int(os.environ.get("ARTESP_GEOJSON_SIMPLIFY_STEP", "2")))
    decimais = max(3, min(6, int(os.environ.get("ARTESP_GEOJSON_DECIMAIS", "4"))))

    def sanitize(o):
        if isinstance(o, dict):
            return {k: sanitize(v) for k, v in o.items()}
        if isinstance(o, list):
            return [sanitize(x) for x in o]
        if isinstance(o, float):
            if math.isnan(o) or math.isinf(o):
                return None
            if -180 <= o <= 180:
                return round(o, decimais)
            return round(o, 6) if abs(o) > 20 else round(o, 3)
        return o

    def process_geojson(o):
        if not isinstance(o, dict):
            return sanitize(o)
        if o.get("type") == "FeatureCollection" and "features" in o:
            out = {}
            for k, v in o.items():
                if k == "features":
                    out[k] = [process_geojson(f) for f in v]
                else:
                    out[k] = sanitize(v)
            return out
        if o.get("type") == "Feature" and "geometry" in o:
            geom = o["geometry"]
            if isinstance(geom, dict) and "coordinates" in geom:
                coords = geom["coordinates"]
                if step > 1:
                    coords = _simplificar_coordenadas(coords, step)
                geom = {**geom, "coordinates": sanitize(coords)}
            else:
                geom = sanitize(geom)
            return {"type": "Feature", "geometry": geom, "properties": sanitize(o.get("properties", {}))}
        return sanitize(o)

    with open(caminho, "w", encoding="utf-8", newline="\n") as f:
        json.dump(process_geojson(obj), f, ensure_ascii=False, separators=(',', ':'), allow_nan=False)


# ══════════════════════════════════════
#  EXTRAIR GEOMETRIA DA MALHA
# ══════════════════════════════════════
def ajustar_offset(lat, lon, local_especifico):
    """Desloca coordenada quando o local representa o 'lado de lá' da rodovia (SUL/OESTE/INTERNA)."""
    txt = str(local_especifico or "").strip().upper()
    # Termos que representam o "lado de lá" (Lote 13: rodovias horizontais = OESTE, INTERNA)
    termos_deslocar = ["SUL", "DEC", "DECRESCENTE", "B", "OESTE", "INTERNA"]
    if any(t in txt for t in termos_deslocar):
        print(f">>> DESLOCANDO: {txt}")
        offset = 0.0005
        return lat + offset, lon + offset
    return lat, lon


def extrair_ponto_geom(rod, km, sentido=None, local=None):
    """Retorna [lon, lat] do ponto mais próximo do km na malha. Offset só se local específico for Sul."""
    pts = CACHE.obter(rod, sentido)
    if not pts:
        return None
    mais_prox = min(pts, key=lambda p: abs(p["km"] - km))
    if abs(mais_prox["km"] - km) > 5.0:
        return None
    lat_f, lon_f = ajustar_offset(mais_prox["lat"], mais_prox["lon"], local or "")
    return [lon_f, lat_f]


def _remover_duplicados_km(pontos, eps_km=0.001):
    """Deduplicação de KM: mantém um ponto por faixa de KM (lista já ordenada por km). Mata zigue-zague de traçados sobrepostos."""
    if not pontos or len(pontos) < 2:
        return pontos
    out = [pontos[0]]
    for p in pontos[1:]:
        if abs(p["km"] - out[-1]["km"]) > eps_km:
            out.append(p)
    return out


def _quebrar_segmentos_distancia(pontos, max_salto_m=MAX_SALTO_M, tolerancia_impossivel_m=TOLERANCIA_QUEBRA_IMPOSSIVEL_M):
    """Quebra de linha por distância impossível ou gap: se dist > (delta_km*1000 + tolerância) ou dist > max_salto_m, inicia novo segmento (MultiLineString)."""
    if not pontos or len(pontos) < 2:
        return [pontos] if pontos else []
    segmentos = [[pontos[0]]]
    for i in range(1, len(pontos)):
        p1, p2 = pontos[i - 1], pontos[i]
        c1 = [p1["lon"], p1["lat"]]
        c2 = [p2["lon"], p2["lat"]]
        dist_m = _distancia_haversine(c1, c2)
        delta_km = abs(p2["km"] - p1["km"])
        limite_impossivel = delta_km * 1000 + tolerancia_impossivel_m
        if dist_m > limite_impossivel or dist_m > max_salto_m:
            segmentos.append([])
        segmentos[-1].append(p2)
    return segmentos


def extrair_linha_geom(rod, km_ini=None, km_fim=None, sentido=None, local=None):
    """Se não achar a Marginal na malha, usa o Eixo da rodovia como base e aplica o deslocamento.
    Usa normalizar_sentido_para_cache(sentido, rod) para traduzir o sentido da planilha para o do Cache (Lote 13: Crescente→Oeste, Decrescente→Leste; ou por rodovia em RODOVIA_SENTIDO_PARA_MALHA).
    1) Tenta buscar com o local usando o sentido traduzido.
    2) SEGUNDA CHANCE: busca pelo eixo geral (local=None) com sentido traduzido.
    3) Processa os pontos aplicando o offset pelo nome do local (SUL/OESTE deslocam).
    """
    sent_busca = normalizar_sentido_para_cache(sentido, rod)
    if km_ini is not None and km_fim is not None:
        # 1. Tenta buscar com o local usando o sentido traduzido para o cache (ex.: Crescente→Oeste)
        pts = CACHE.obter_intervalo(rod, km_ini, km_fim, sent_busca, local)
        # Trava só quando o cache TEM a chave específica mas pts vazio (evita duplicidade em malhas por local)
        # Se o cache NÃO tem a chave específica (L26: malha só por rodovia), usa fallback local=None
        if local and (not pts or len(pts) < 2):
            chave_especifica = CACHE._chave(rod, sent_busca, local)
            tem_chave = chave_especifica in CACHE.dados and CACHE.dados.get(chave_especifica)
            if tem_chave:
                print(f"AVISO: {chave_especifica} existe mas sem pts no intervalo. Pulando.")
                return None
        # 2. SEGUNDA CHANCE: eixo geral (local=None) com sentido traduzido
        if not pts or len(pts) < 2:
            pts = CACHE.obter_intervalo(rod, km_ini, km_fim, sent_busca, None)
        if not pts or len(pts) < 2:
            p1 = extrair_ponto_geom(rod, km_ini, sent_busca, local)
            p2 = extrair_ponto_geom(rod, km_fim, sent_busca, local)
            if p1 and p2 and p1 != p2:
                return {"type": "LineString", "coordinates": [p1, p2]}
            return None
    else:
        pts = CACHE.obter(rod, sent_busca, local)
        if local and (not pts or len(pts) < 2):
            chave_especifica = CACHE._chave(rod, sent_busca, local)
            tem_chave = chave_especifica in CACHE.dados and CACHE.dados.get(chave_especifica)
            if tem_chave:
                print(f"AVISO: {chave_especifica} existe mas sem pts. Pulando.")
                return None
        if not pts or len(pts) < 2:
            pts = CACHE.obter(rod, sent_busca, None)
        if not pts or len(pts) < 2:
            return None

    # 1. Ordenação já feita no cache (Rodovia > Sentido > KM). 2. Deduplicação de KM (evita zigue-zague)
    pts = _remover_duplicados_km(pts)
    # 3. Quebra por distância impossível ou gap (L21: não ligar se dist > 100m; criar novo segmento)
    segmentos_pts = _quebrar_segmentos_distancia(pts, MAX_SALTO_M, TOLERANCIA_QUEBRA_IMPOSSIVEL_M)

    # Diagnóstico: descomente para ver quantos pontos por rodovia/sentido/local.
    # print(f"Rodovia: {rod} | Sentido: {sentido} | Local: {local} | Pontos: {len(pts)} | Segmentos: {len(segmentos_pts)}")

    # 3. Processa os pontos do eixo aplicando o offset baseado no nome do local (SUL/OESTE deslocam)
    linhas_coords = []
    for seg in segmentos_pts:
        if not seg or len(seg) < 2:
            continue
        coords = []
        for p in seg:
            n_lat, n_lon = ajustar_offset(p["lat"], p["lon"], local or "")
            coords.append([n_lon, n_lat])
        coords = _dedup(coords)
        if len(coords) < 2:
            continue
        coords = _filtrar_salto_geografico(coords)
        coords = _filtrar_mudanca_brusca(coords)
        coords = filtrar_espaco_minimo(coords)
        coords = simplificar(coords)
        coords = _limitar_pontos_linha(coords)
        if len(coords) >= 2:
            linhas_coords.append(coords)

    if not linhas_coords:
        return None
    if len(linhas_coords) == 1:
        return {"type": "LineString", "coordinates": linhas_coords[0]}
    return {"type": "MultiLineString", "coordinates": linhas_coords}


def extrair_sentido(locais):
    """Igual ao script mãe: extrai sentido a partir da lista de locais. Retorna 'N/S', 'N', etc."""
    if not locais:
        return "-"
    locais_list = locais if isinstance(locais, (list, tuple)) else [locais]
    sentidos = []
    for loc in locais_list:
        loc_str = str(loc).strip() if loc else ""
        if loc_str and loc_str in MAPA_SENTIDO:
            s = MAPA_SENTIDO[loc_str]
            if s not in sentidos:
                sentidos.append(s)
    return "/".join(sorted(sentidos)) if sentidos else "-"


def _extrair_pontos_interesse_geometria(geom, ki, kf):
    """
    Igual ao script mãe: retorna lista de {"coordinates": [lon, lat], "km": float, "tipo": "inicial"|"final"|"ponto"}.
    """
    if not geom or not isinstance(geom, dict):
        return []
    gtype = geom.get("type") or ""
    coords = geom.get("coordinates")
    ki_f = float(ki) if ki is not None else 0.0
    kf_f = float(kf) if kf is not None else 0.0
    out = []

    if gtype == "Point" and isinstance(coords, (list, tuple)) and len(coords) >= 2:
        out.append({"coordinates": list(coords)[:2], "km": _snap_km(ki_f), "tipo": "ponto"})
    elif gtype == "MultiPoint" and isinstance(coords, list):
        n = len(coords)
        for i, c in enumerate(coords):
            if isinstance(c, (list, tuple)) and len(c) >= 2:
                km = _snap_km(ki_f + (kf_f - ki_f) * (i / (n - 1))) if n > 1 else _snap_km(ki_f)
                out.append({"coordinates": list(c)[:2], "km": km, "tipo": "ponto"})
    elif gtype == "LineString" and isinstance(coords, list) and len(coords) >= 2:
        out.append({"coordinates": list(coords[0])[:2], "km": _snap_km(ki_f), "tipo": "inicial"})
        out.append({"coordinates": list(coords[-1])[:2], "km": _snap_km(kf_f), "tipo": "final"})
    elif gtype == "MultiLineString" and isinstance(coords, list):
        if coords:
            first_line = coords[0]
            last_line = coords[-1]
            if isinstance(first_line, (list, tuple)) and len(first_line) >= 1:
                c0 = first_line[0]
                if isinstance(c0, (list, tuple)) and len(c0) >= 2:
                    out.append({"coordinates": list(c0)[:2], "km": _snap_km(ki_f), "tipo": "inicial"})
            if isinstance(last_line, (list, tuple)) and len(last_line) >= 1:
                c1 = last_line[-1]
                if isinstance(c1, (list, tuple)) and len(c1) >= 2:
                    out.append({"coordinates": list(c1)[:2], "km": _snap_km(kf_f), "tipo": "final"})
    return out


def _expandir_features_com_marcadores_alfinete(features_list):
    """
    Igual ao script mãe: para cada feature com pontos_interesse, adiciona features Point
    com tipo_marcador (inicial/final/ponto) e cor_hex. pontos_interesse = [{"coordinates": [...], "km": x, "tipo": "..."}, ...].
    """
    out = []
    for ft in features_list:
        out.append(ft)
        p = ft.get("properties") or {}
        if p.get("marcador_apenas"):
            continue
        pontos = p.get("pontos_interesse") or []
        fid = p.get("id") or ""
        for i, pt in enumerate(pontos):
            coords = pt.get("coordinates") if isinstance(pt, dict) else (pt if isinstance(pt, (list, tuple)) and len(pt) >= 2 else None)
            tipo = pt.get("tipo", "ponto") if isinstance(pt, dict) else "ponto"
            if not isinstance(coords, (list, tuple)) or len(coords) < 2:
                continue
            cor = CORES_MARCADORES.get(tipo, CORES_MARCADORES.get("ponto", "#D32F2F"))
            id_marcador = f"{fid}_m{tipo[:1]}{i}"[:50]
            km_val = pt.get("km") if isinstance(pt, dict) else None
            props_marcador = dict(p)
            props_marcador.update({
                "id": id_marcador,
                "marcador_apenas": True,
                "tipo_marcador": tipo,
                "cor_hex": cor,
                "id_feature": fid,
                "km_inicial": km_val if km_val is not None else p.get("km_inicial"),
                "km_final": km_val if km_val is not None else p.get("km_final"),
            })
            out.append({
                "type": "Feature",
                "geometry": {"type": "Point", "coordinates": [float(coords[0]), float(coords[1])]},
                "properties": props_marcador
            })
    return out


# ══════════════════════════════════════
#  CARREGAR MALHA / EIXO
# ══════════════════════════════════════
def carregar_malha(caminho, lote_sigla):
    """Carrega CSV/XLSX/GeoJSON de eixo para o CACHE."""
    global DADOS_MALHA, CACHE

    CACHE.limpar()
    DADOS_MALHA.clear()

    if not os.path.exists(caminho):
        print(f"[ERRO] Malha não encontrada: {caminho}")
        return

    ext = os.path.splitext(caminho)[1].lower()

    if ext == ".geojson":
        _carregar_malha_geojson(caminho, lote_sigla)
    elif ext in (".csv", ".xlsx", ".xls"):
        _carregar_malha_tabular(caminho, lote_sigla)
    else:
        print(f"[AVISO] Extensão não suportada: {ext}")

    if CORRIGIR_EIXO:
        try:
            _aplicar_correcao_eixo_ao_cache()
        except Exception as e:
            print(f"  [AVISO] Correção de eixo ignorada (erro: {e}). Malha usada sem correção.")
            import traceback
            traceback.print_exc()

    CACHE.finalizar_carregamento()
    n_pontos = sum(len(v) for v in CACHE.dados.values())
    n_rods = len(CACHE._rodovias)
    print(f"  [OK] Malha: {n_pontos} pontos, {n_rods} rodovias, sentidos={'SIM' if CACHE.tem_sentidos else 'NÃO'}")


def _normalizar_sentido_malha(s_raw, lote_sigla=None):
    """Normaliza valor da coluna Sentido da malha para Crescente/Decrescente (cache unificado).
    L21: malha pode vir com Norte/Sul/Leste/Oeste ou Externa/Interna — mapear para Crescente/Decrescente."""
    s = (s_raw or "").strip()
    if not s:
        return None
    low = s.lower()
    if low in ("crescente", "cresc"):
        return "Crescente"
    if low in ("decrescente", "decresc"):
        return "Decrescente"
    if lote_sigla in ("L21", "L26"):
        if low in ("externo", "externa"):
            return "Crescente"
        if low in ("interno", "interna"):
            return "Decrescente"
        # Malha L21 pode ter coluna Sentido = Norte, Sul, Leste, Oeste — unificar para cache
        mapa = LOTE_SENTIDO_PARA_MALHA.get(lote_sigla, {})
        if low in ("norte", "n"):
            return mapa.get("Norte", "Decrescente")
        if low in ("sul", "s"):
            return mapa.get("Sul", "Crescente")
        if low in ("leste", "l"):
            return mapa.get("Leste", "Crescente")
        if low in ("oeste", "o"):
            return mapa.get("Oeste", "Decrescente")
    return s


def normalizar_sentido_para_cache(sentido_planilha, rodovia=None):
    """Traduz o sentido da planilha/app para o sentido com que a malha está no Cache.
    L21/L26: cache usa Crescente/Decrescente — retornar igual. L13: Crescente→Oeste, Decrescente→Leste.
    Quando sentido é None/vazio (ex.: L13 com geometria por rodovia), retorna None para agregar todos os sentidos no cache.
    """
    s = str(sentido_planilha or "").strip().upper()
    if not s or s in ("NONE", "NAN"):
        return None
    if s in ("CRESCENTE", "CRESC", "NORTE", "A"):
        snorm = "Crescente"
    elif s in ("DECRESCENTE", "DECRESC", "SUL", "B"):
        snorm = "Decrescente"
    else:
        snorm = None

    # L21 e L26: malha usa Crescente/Decrescente — retornar direto
    rodovias_l21 = set(RODOVIAS_POR_LOTE.get("L21", []))
    rodovias_l26 = set(RODOVIAS_POR_LOTE.get("L26", []))
    if rodovia and snorm and (rodovia in rodovias_l21 or rodovia in rodovias_l26):
        return snorm

    if rodovia and RODOVIA_SENTIDO_PARA_MALHA.get(rodovia):
        regra = RODOVIA_SENTIDO_PARA_MALHA[rodovia]
        # Cache é preenchido com sentido normalizado (Crescente/Decrescente) via _normalizar_sentido_malha;
        # retornar snorm para bater com a chave do cache, não chave_malha (Norte/Sul).
        for chave_malha, valor in regra.items():
            if valor == snorm:
                return snorm
        if snorm is None and s in ("LESTE", "OESTE", "NORTE", "SUL"):
            return s.title()

    # Padrão Lote 13: malha com Leste/Oeste; aceita também Leste/Oeste/Externa/Interna direto
    traducao = {
        "CRESCENTE": "Oeste",
        "DECRESCENTE": "Leste",
        "NORTE": "Leste",
        "SUL": "Oeste",
        "LESTE": "Leste",
        "OESTE": "Oeste",
        "EXTERNA": "Externa",
        "INTERNA": "Interna",
        "A": "Oeste",
        "B": "Leste",
    }
    return traducao.get(s, sentido_planilha if sentido_planilha is not None else s)


def descobrir_sentido_malha(nome_local, lote_sigla, rodovia=None):
    """Traduz Local (ex: PISTA_SUL) para sentido da malha (ex: Decrescente).
    Usa RODOVIA_SENTIDO_PARA_MALHA quando rodovia é informada; senão LOTE_SENTIDO_PARA_MALHA.
    """
    nome = (nome_local or "").strip().upper()
    if not nome:
        return "Crescente"
    regra = (RODOVIA_SENTIDO_PARA_MALHA.get(rodovia or "", {}) or
             LOTE_SENTIDO_PARA_MALHA.get(lote_sigla, {}))
    if "NORTE" in nome:
        return regra.get("Norte", "Crescente")
    if "SUL" in nome:
        return regra.get("Sul", "Decrescente")
    if "LESTE" in nome:
        return regra.get("Leste", regra.get("Norte", "Crescente"))
    if "OESTE" in nome:
        return regra.get("Oeste", regra.get("Sul", "Decrescente"))
    if "EXTERNA" in nome or "CRESC" in nome:
        return regra.get("Externa", regra.get("Crescente", "Crescente"))
    if "INTERNA" in nome or "DECRESC" in nome:
        return regra.get("Interna", regra.get("Decrescente", "Decrescente"))
    return "Crescente"


def _carregar_malha_tabular(caminho, lote_sigla):
    """Carrega malha de CSV ou Excel. Padronização: A=Rodovia, B=Km, C=Sentido, D=Latitude, E=Longitude (F=Local opcional)."""
    try:
        if caminho.endswith(".csv"):
            df = pd.read_csv(caminho, sep=None, engine="python", encoding="utf-8-sig")
        else:
            df = pd.read_excel(caminho)
    except Exception as e:
        print(f"[ERRO] Leitura malha: {e}")
        return

    df.columns = [str(c).strip() for c in df.columns]

    # Detectar colunas — padronização: A=Rodovia, B=Km, C=Sentido, D=Latitude, E=Longitude
    # Suporte a malha com dois sentidos: LAT_NORTE/LON_NORTE e LAT_SUL/LON_SUL (chave = KM + Sentido)
    col_map = {}
    for c in df.columns:
        ck = norm_key(c)
        if "rodovia" in ck or "rod" == ck:
            col_map["rodovia"] = c
        elif "km" in ck and "interno" not in ck:
            col_map["km"] = c
        elif "kminterno" in ck:
            col_map["km_interno"] = c
        elif "sentido" in ck or "sent" == ck:
            col_map["sentido"] = c
        elif "local" in ck and "para" not in ck:
            col_map["local"] = c
        elif "norte" in ck and ("lat" in ck or ck == "latnorte"):
            col_map["lat_norte"] = c
        elif "norte" in ck and ("lon" in ck or "lng" in ck or ck == "lonnorte"):
            col_map["lon_norte"] = c
        elif "sul" in ck and ("lat" in ck or ck == "latsul"):
            col_map["lat_sul"] = c
        elif "sul" in ck and ("lon" in ck or "lng" in ck or ck == "lonsul"):
            col_map["lon_sul"] = c
        elif "latitude" in ck or ck == "lat":
            col_map.setdefault("lat", c)
        elif "longitude" in ck or ck in ("lon", "lng", "long"):
            col_map.setdefault("lon", c)

    if "km" not in col_map and "km_interno" in col_map:
        col_map["km"] = col_map["km_interno"]

    tem_par_norte = "lat_norte" in col_map and "lon_norte" in col_map
    tem_par_sul = "lat_sul" in col_map and "lon_sul" in col_map
    tem_par_unico = "lat" in col_map and "lon" in col_map
    if not (tem_par_unico or tem_par_norte or tem_par_sul) or "rodovia" not in col_map or "km" not in col_map:
        print(f"[ERRO] Malha: precisa de Rodovia, Km e (Lat/Lon ou LAT_NORTE/LON_NORTE e/ou LAT_SUL/LON_SUL). Colunas: {list(df.columns)}")
        return

    mapa_sentido_lote = LOTE_SENTIDO_PARA_MALHA.get(lote_sigla, {})

    for _, row in df.iterrows():
        rod = normalizar_rodovia(row.get(col_map["rodovia"]))
        if not rod:
            continue

        km = _parse_km_excel(row.get(col_map["km"]))
        if km is None:
            continue

        # Por rodovia tem prioridade sobre o do lote (ex.: SP0000127 Norte=Decrescente)
        mapa_rod = RODOVIA_SENTIDO_PARA_MALHA.get(rod, {}) or mapa_sentido_lote

        lat_n, lon_n = None, None
        sentido = None
        if "sentido" in col_map:
            s_raw = str(row.get(col_map["sentido"], "")).strip()
            if s_raw and s_raw.lower() not in ("", "nan", "none"):
                sentido = _normalizar_sentido_malha(s_raw, lote_sigla)

        def _adicionar_ponto(lat_val, lon_val, sentido_malha, local_malha=None):
            if lat_val is None or lon_val is None:
                return
            # Chave do cache inclui sentido para não misturar coordenadas Norte/Sul na geometria
            CACHE.adicionar(rod, sentido_malha, km, lon_val, lat_val, local=local_malha)
            DADOS_MALHA[CACHE._chave(rod, sentido_malha, local_malha)] = True

        # Malha com LAT_NORTE/LON_NORTE e LAT_SUL/LON_SUL: IF explícito — cada sentido usa SEU par de coordenadas
        sent_norte = sentido or mapa_rod.get("Norte", "Crescente")
        sent_sul = sentido or mapa_rod.get("Sul", "Decrescente")
        if tem_par_norte:
            lat_n = _parse_coord_malha(row.get(col_map["lat_norte"]))
            lon_n = _parse_coord_malha(row.get(col_map["lon_norte"]))
            if lat_n is not None and lon_n is not None:
                _adicionar_ponto(lat_n, lon_n, sent_norte, None)

        if tem_par_sul:
            lat_s = _parse_coord_malha(row.get(col_map["lat_sul"]))
            lon_s = _parse_coord_malha(row.get(col_map["lon_sul"]))
            if lat_s is not None and lon_s is not None:
                _adicionar_ponto(lat_s, lon_s, sent_sul, None)
            elif tem_par_norte and lat_n is not None and lon_n is not None:
                # Sul vazio: usar coords Norte; o offset é aplicado na hora de desenhar (extrair_linha_geom)
                _adicionar_ponto(lat_n, lon_n, sent_sul, None)

        if tem_par_unico and not tem_par_norte and not tem_par_sul:
            lat = _parse_coord_malha(row.get(col_map["lat"]))
            lon = _parse_coord_malha(row.get(col_map["lon"]))
            if lat is None or lon is None:
                continue
            locais_linha = []
            if "local" in col_map:
                coluna_local = row.get(col_map["local"], "")
                l_raw = str(coluna_local).strip() if coluna_local is not None else ""
                if l_raw and l_raw.upper() not in ("", "NAN", "NONE"):
                    for parte in l_raw.replace(",", ";").replace("\t", ";").split(";"):
                        nome_local = parte.strip().upper().replace(" ", "_")
                        if nome_local and nome_local not in ("NAN", "NONE"):
                            locais_linha.append(nome_local)
            if locais_linha:
                for local_malha in locais_linha:
                    sentido_linha = sentido if sentido is not None else descobrir_sentido_malha(local_malha, lote_sigla, rod)
                    _adicionar_ponto(lat, lon, sentido_linha, local_malha)
            else:
                sentido_linha = sentido or "Crescente"
                _adicionar_ponto(lat, lon, sentido_linha, None)


def _carregar_malha_geojson(caminho, lote_sigla):
    """Carrega malha de GeoJSON."""
    try:
        with open(caminho, "r", encoding="utf-8") as f:
            gj = json.load(f)
    except Exception as e:
        print(f"[ERRO] GeoJSON malha: {e}")
        return

    for feat in gj.get("features", []):
        props = feat.get("properties", {})
        rod = normalizar_rodovia(props.get("Rodovia") or props.get("rodovia") or "")
        if not rod:
            continue

        s_raw = props.get("Sentido") or props.get("sentido") or None
        sentido = _normalizar_sentido_malha(str(s_raw), lote_sigla) if s_raw is not None else None
        l_raw = props.get("Local") or props.get("local") or None
        local_malha = str(l_raw).strip().upper().replace(" ", "_") if l_raw else None
        if local_malha and sentido is None:
            sentido = descobrir_sentido_malha(local_malha, lote_sigla, rod)
        if sentido is None:
            sentido = "Crescente"

        geom = feat.get("geometry", {})
        coords = geom.get("coordinates", [])
        g_type = geom.get("type", "")

        if g_type == "LineString":
            for i, pt in enumerate(coords):
                if len(pt) >= 2:
                    km = props.get("Km", 0) + i * 0.01
                    CACHE.adicionar(rod, sentido, km, pt[0], pt[1], local=local_malha)
        elif g_type == "Point" and len(coords) >= 2:
            km = _parse_km_excel(props.get("Km") or props.get("km") or 0)
            CACHE.adicionar(rod, sentido, km or 0, coords[0], coords[1], local=local_malha)

        DADOS_MALHA[CACHE._chave(rod, sentido, local_malha)] = True


# ══════════════════════════════════════
#  APLICAR KM OFICIAL
# ══════════════════════════════════════
def aplicar_km_oficial_malha_em_todas_rodovias_v2(df_eixo, df_malha):
    """Converte Km interno do eixo para KM oficial da concessão (MALHA). v2 com fallback de sentido."""
    M = df_malha.copy()
    M.columns = [str(c).strip().lower() for c in M.columns]

    if "rodovia" not in M.columns or "km" not in M.columns:
        return df_eixo

    E = df_eixo.copy()
    E["Rodovia_norm"] = E["Rodovia"].apply(normalizar_rodovia) if "Rodovia" in E.columns else ""

    if "Km_interno" not in E.columns:
        E["Km_interno"] = E.get("Km", 0)

    for rod in E["Rodovia_norm"].unique():
        if not rod:
            continue

        m_rod = M[M["rodovia"].apply(normalizar_rodovia) == rod]
        if m_rod.empty:
            continue

        km_min_m = float(m_rod["km"].min()) * 1000
        km_max_m = float(m_rod["km"].max()) * 1000
        span_m = km_max_m - km_min_m

        if span_m <= 0:
            continue

        sentidos_eixo = set(E.loc[E["Rodovia_norm"] == rod, "Sentido"].astype(str).str.strip().unique()) if "Sentido" in E.columns else set()
        sentidos_eixo.discard("")
        sentidos_eixo.discard("nan")

        sent_cresc = "Crescente"
        sent_decr = "Decrescente"

        if sent_cresc not in sentidos_eixo or sent_decr not in sentidos_eixo:
            sentidos_ord = sorted(sentidos_eixo)
            fb_cresc = sentidos_ord[0] if len(sentidos_ord) >= 1 else None
            fb_decr = sentidos_ord[1] if len(sentidos_ord) >= 2 else None
        else:
            fb_cresc, fb_decr = sent_cresc, sent_decr

        for sent in sentidos_eixo:
            idx = E.index[(E["Rodovia_norm"] == rod) & (E["Sentido"].astype(str).str.strip() == sent)]
            if len(idx) == 0:
                continue

            sub = E.loc[idx].sort_values("Km_interno")
            s0 = float(sub["Km_interno"].min())
            s1 = float(sub["Km_interno"].max())
            L = s1 - s0
            if L <= 0:
                continue

            frac = (sub["Km_interno"] - s0) / L

            if fb_cresc is not None and sent == fb_cresc:
                km_new_m = km_min_m + frac * span_m
            elif fb_decr is not None and sent == fb_decr:
                km_new_m = km_max_m - frac * span_m
            else:
                km_new_m = km_min_m + frac * span_m

            E.loc[sub.index, "Km"] = km_new_m / 1000.0

    return E


# ══════════════════════════════════════
#  LEITURA DO EXCEL
# ══════════════════════════════════════
def ler_excel(caminho, sheet=0, cabecalho_linhas=CABECALHO_LINHAS, linha_inicio_dados=LINHA_INICIO_DADOS):
    """Lê o Excel: linha 1 = cabeçalho, linhas 2-5 = exemplos/instruções. Dados a partir da linha 6."""

    def _ler(path):
        df_raw = pd.read_excel(path, sheet_name=sheet, header=None)
        if len(df_raw) == 0:
            return pd.DataFrame(), [], path
        cabecalho = df_raw.iloc[:cabecalho_linhas].values.tolist()
        col_names = df_raw.iloc[0].values
        df = df_raw.iloc[linha_inicio_dados - 1:].copy()
        df.columns = col_names
        df = df.reset_index(drop=True)
        return df, cabecalho

    try:
        df, cabecalho = _ler(caminho)
        return df, cabecalho, caminho
    except Exception:
        tmp = os.path.join(
            tempfile.gettempdir(),
            f"artesp_{datetime.datetime.now().strftime('%H%M%S')}.xlsx"
        )
        shutil.copy2(caminho, tmp)
        try:
            df, cabecalho = _ler(tmp)
            return df, cabecalho, tmp
        except Exception:
            raise


def _corrigir_lat_lon_trocadas(df):
    if df is None or df.empty or "Latitude" not in df.columns or "Longitude" not in df.columns:
        return df
    try:
        lat_vals = pd.to_numeric(df["Latitude"], errors="coerce").dropna()
        lon_vals = pd.to_numeric(df["Longitude"], errors="coerce").dropna()
        if lat_vals.empty or lon_vals.empty:
            return df
        med_lat = float(lat_vals.median())
        med_lon = float(lon_vals.median())
        if (-54 <= med_lat <= -40 and -26 <= med_lon <= -18):
            df = df.copy()
            df["Latitude"], df["Longitude"] = df["Longitude"].copy(), df["Latitude"].copy()
            print("  [OK] Latitude/Longitude corrigidas (estavam trocadas).")
    except Exception:
        pass
    return df


# ══════════════════════════════════════
#  NORMALIZAR COLUNAS DO DATAFRAME
# ══════════════════════════════════════
def normalizar_colunas_df(df):
    """Renomeia colunas do Excel para nomes padronizados internos."""
    if df is None or df.empty:
        return df

    mapa_colunas = {}
    for c in df.columns:
        ck = norm_key(str(c))
        if "lote" in ck:
            mapa_colunas[c] = "lote"
        elif "rodovia" in ck or ck == "rod":
            mapa_colunas[c] = "rodovia"
        elif "programa" in ck:
            mapa_colunas[c] = "programa"
        elif "subitem" in ck:
            mapa_colunas[c] = "subitem"
        elif "item" in ck and "sub" not in ck and "detalh" not in ck:
            mapa_colunas[c] = "item"
        elif "detalhamento" in ck or "descricao" in ck or "servico" in ck:
            mapa_colunas[c] = "detalhamento_servico"
        elif "unidade" in ck or "unid" in ck or ck == "un":
            mapa_colunas[c] = "unidade"
        elif "quantidade" in ck or "qtd" in ck or "qtde" in ck:
            mapa_colunas[c] = "quantidade"
        elif "kminicial" in ck or "kminicio" in ck or "kmini" in ck:
            mapa_colunas[c] = "km_inicial"
        elif "kmfinal" in ck or "kmfim" in ck:
            mapa_colunas[c] = "km_final"
        elif "local" in ck or "pista" in ck:
            mapa_colunas[c] = "local"
        elif "datainicial" in ck or "datainicio" in ck or "dataini" in ck:
            mapa_colunas[c] = "data_inicial"
        elif "datafinal" in ck or "datafim" in ck:
            mapa_colunas[c] = "data_final"
        elif "observ" in ck or "obs" in ck:
            mapa_colunas[c] = "observacoes_gerais"
        elif "latitude" in ck or "lat" == ck:
            mapa_colunas[c] = "Latitude"
        elif "longitude" in ck or "lon" == ck or "lng" == ck:
            mapa_colunas[c] = "Longitude"

    df = df.rename(columns=mapa_colunas)
    return df


# ══════════════════════════════════════
#  FILTRAR DADOS POR LOTE / RECONHECER LOTE NO EXCEL
# ══════════════════════════════════════
def detectar_lote_no_excel(df, lotes_disponiveis=None):
    """
    Reconhece o lote a partir da coluna 'lote' do DataFrame (já normalizado).
    Retorna (lote_key, mensagem):
      - (str, None) se um único lote válido for encontrado (ex.: "13" para L13)
      - (None, str) se nenhuma coluna lote, múltiplos lotes ou lote não reconhecido
    lotes_disponiveis: dict opcional { "13": {...}, "21": {...} }; se None, usa LOTES.
    """
    if df is None or df.empty:
        return None, "Planilha vazia ou inválida."
    if "lote" not in df.columns:
        return None, "Coluna 'lote' não encontrada na planilha."
    lotes = lotes_disponiveis if lotes_disponiveis is not None else LOTES
    siglas_validas = {info["sigla"] for info in lotes.values() if isinstance(info, dict) and info.get("sigla")}
    valores = df["lote"].dropna().astype(str).str.strip()
    valores = valores[valores.str.len() > 0]
    if valores.empty:
        return None, "Nenhum valor de lote preenchido na planilha."
    normalizados = valores.apply(normalizar_lote).dropna()
    unicos = normalizados.unique().tolist()
    # Só aceitar siglas que existem em LOTES (L13, L21, L26)
    unicos_validos = [u for u in unicos if u in siglas_validas]
    if len(unicos_validos) == 0:
        return None, f"Lote(s) na planilha não reconhecido(s): {', '.join(unicos)}. Use L13, L21 ou L26."
    if len(unicos_validos) > 1:
        # Planilha mãe tem um único lote: usar o que aparece mais vezes (ignora exemplo/avulso)
        contagem = normalizados.value_counts()
        validos_contagem = [(s, contagem.get(s, 0)) for s in unicos_validos]
        validos_contagem.sort(key=lambda x: -x[1])
        sigla = validos_contagem[0][0]
    else:
        sigla = unicos_validos[0]
    lote_key = next((k for k, v in lotes.items() if isinstance(v, dict) and v.get("sigla") == sigla), None)
    return (lote_key, None)


def filtrar_dados_por_lote(df, lote_sigla):
    """Filtra linhas que pertencem ao lote selecionado."""
    if "lote" not in df.columns:
        return df
    mask = df["lote"].apply(lambda x: normalizar_lote(x) == lote_sigla)
    return df[mask].copy()


# ══════════════════════════════════════
#  NOMES DE ARQUIVOS
# ══════════════════════════════════════
def template_filename(lote_sigla, chave_modalidade):
    """Nome do arquivo template Excel."""
    return f"{lote_sigla}_{chave_modalidade}_template.xlsx"


def _path_asset_template(lote_sigla, chave_modalidade, ano=None, versao=None):
    """Resolve template Excel: tenta template_filename, depois {lote}_{mod}_{ano}_{versao}.xlsx.
    Para conservação aceita também nome com 'conserva' (L13_conserva_2026_r0.xlsx)."""
    nome_padrao = template_filename(lote_sigla, chave_modalidade)
    path = _path_asset("template", nome_padrao)
    if os.path.exists(path):
        return path
    chaves_tentar = [chave_modalidade]
    if (chave_modalidade or "").strip().lower() == "conservacao":
        chaves_tentar.append("conserva")
    for ano_try in (ano, 2026, 2025) if ano else (2026, 2025):
        for ver_try in (versao, "r0") if versao else ("r0",):
            for chave in chaves_tentar:
                nome_alt = f"{lote_sigla}_{chave}_{ano_try}_{ver_try}.xlsx"
                candidato = _path_asset("template", nome_alt)
                if os.path.exists(candidato):
                    return candidato
    return path


def _versao_suffix_arquivo(ver):
    s = str(ver or "").strip()
    return s.upper() if s else s


def basename_saida(lote_sigla, chave_modalidade, ano, versao, mes=None, tipo=None):
    """
    Gera nome base para arquivos de saída no formato exato:
    - Anual: L13_conservacao_2026_R0
    - Executado (mês anterior): L26_conservacao_executado_janeiro_2026_R02
    - Programado/Mensal (mês seguinte): L13_conservacao_programado_março_2026_R01
    """
    mod_saida = "conservacao" if (chave_modalidade or "").strip().lower() == "conserva" else (chave_modalidade or "")
    if not mes or tipo not in ("MENSAL", "EXECUTADO"):
        return f"{lote_sigla}_{mod_saida}_{ano}_{_versao_suffix_arquivo(versao)}"
    mes_nome = MESES_NOME_SAIDA.get(int(mes), f"{int(mes):02d}")
    if tipo == "EXECUTADO":
        ver = "r02" if versao == "e" else versao
        return f"{lote_sigla}_{mod_saida}_executado_{mes_nome}_{ano}_{_versao_suffix_arquivo(ver)}"
    if tipo == "MENSAL":
        ver = "r01" if versao == "m" else versao
        return f"{lote_sigla}_{mod_saida}_programado_{mes_nome}_{ano}_{_versao_suffix_arquivo(ver)}"
    return f"{lote_sigla}_{mod_saida}_{ano}_{_versao_suffix_arquivo(versao)}_{mes:02d}"


def criar_saida(base_dir, ano, lote_sigla, tipo=None, subpasta_mes=None):
    """Cria e retorna pasta de saída."""
    pasta = os.path.join(base_dir, str(ano), lote_sigla)
    if subpasta_mes:
        pasta = os.path.join(pasta, subpasta_mes)
    os.makedirs(pasta, exist_ok=True)
    return pasta


# ══════════════════════════════════════
#  ESCOLHER PERÍODO
# ══════════════════════════════════════
def escolher_periodo(tipo):
    """Calcula período (mês) automaticamente conforme tipo."""
    hoje = datetime.date.today()
    if tipo == "MENSAL":
        nxt = hoje + relativedelta(months=1)
        return {"ano": nxt.year, "mes": nxt.month, "tipo": tipo}
    elif tipo == "EXECUTADO":
        prv = hoje - relativedelta(months=1)
        return {"ano": prv.year, "mes": prv.month, "tipo": tipo}
    return {"ano": hoje.year, "mes": None, "tipo": tipo}


# ═══════════════════════════════════════════════════════════════
#  [P1] PERÍODO — função pura, sem tkinter, importável pelo backend web
# ═══════════════════════════════════════════════════════════════
def calcular_periodo_mensal(versao: str, data_referencia=None):
    """
    Calcula (ano, mes) para relatórios mensais.
    versao: 'M' (programado = mês seguinte) ou 'E' (executado = mês anterior).
    data_referencia: date opcional (padrão: hoje).
    Retorna (ano, mes) ou (None, None) se versão não se aplica.

    Pode ser importado por qualquer módulo sem dependência de GUI.
    """
    v = (versao or "").strip().upper()
    if v not in ("M", "E"):
        return None, None

    ref = data_referencia or datetime.date.today()

    try:
        from dateutil.relativedelta import relativedelta
        if v == "M":
            alvo = ref + relativedelta(months=1)
        else:
            alvo = ref - relativedelta(months=1)
        return alvo.year, alvo.month
    except ImportError:
        if v == "M":
            if ref.month == 12:
                return ref.year + 1, 1
            return ref.year, ref.month + 1
        else:
            if ref.month == 1:
                return ref.year - 1, 12
            return ref.year, ref.month - 1


# ══════════════════════════════════════
#  VALIDAÇÃO JSON SCHEMA
# ══════════════════════════════════════
def validar_json(obj, schema_path, lote):
    r = {"ok": False, "msg": "", "sha": sha256_arquivo(schema_path)}
    if not JSONSCHEMA_OK:
        r["msg"] = "jsonschema nao disponivel"
        return r
    try:
        with open(schema_path, "r", encoding="utf-8") as f:
            schema = json.load(f)

        if lote and re.match(r"^L\d{2}$", lote):
            for def_name in ("FeatureConservacao", "FeatureObras"):
                try:
                    feat = schema.get("$defs", {}).get(def_name)
                    if feat and "properties" in feat:
                        props = feat["properties"].get("properties", {})
                        if isinstance(props, dict) and "properties" in props:
                            lote_prop = props["properties"].get("lote", {})
                            if isinstance(lote_prop.get("enum"), list) and lote not in lote_prop["enum"]:
                                lote_prop["enum"].append(lote)
                except (KeyError, TypeError):
                    pass

        def mult_tol(v, m, i, s):
            if not isinstance(i, (int, float)):
                return
            r = i % m
            if r > 1e-9 and abs(r - m) > 1e-9:
                yield jsonschema.ValidationError(f"{i} not multiple of {m}")

        Val = jsonschema.validators.extend(jsonschema.validators.validator_for(schema), {"multipleOf": mult_tol})
        erros = list(Val(schema).iter_errors(obj))
        if not erros:
            r["ok"] = True
            r["msg"] = "VALIDACAO: APROVADO"
            return r
        msg = [f"VALIDACAO FALHOU: {len(erros)} erros"]
        for i, e in enumerate(erros[:10]):
            msg.append(f"  {i + 1}. {e.message} (em {list(e.path)})")
        if len(erros) > 10:
            msg.append(f"  ... +{len(erros) - 10} erros")
        r["msg"] = "\n".join(msg)
        return r
    except Exception as e:
        r["msg"] = f"ERRO: {e}"
        return r


# ══════════════════════════════════════
#  GERAR ID
# ══════════════════════════════════════
def gerar_id(lote, rod, item, ki, kf, sentido, seq):
    rod_n = normalizar_rodovia(rod)[-10:]
    item_n = normalizar_item(item)[:8]
    ki_f = _snap_km(_parse_km_excel(ki)) or 0
    kf_f = _snap_km(_parse_km_excel(kf)) or 0
    sent = sentido[:1] if sentido else "X"
    id_str = f"{lote}_{rod_n}_{item_n}_{ki_f:.1f}_{kf_f:.1f}_{sent}_{seq:04d}"
    if len(id_str) > 50:
        id_str = f"{lote}_{seq:06d}_{sent}_{hashlib.md5(id_str.encode()).hexdigest()[:6]}"
    return id_str[:50]


# ══════════════════════════════════════
#  RESUMO POR RODOVIA/SENTIDO
# ══════════════════════════════════════
def gerar_resumo_rodovia_sentido(features):
    """Gera resumo de contagem de features por rodovia/sentido."""
    resumo = OrderedDict()
    for ft in features:
        props = ft.get("properties", {})
        rod = props.get("rodovia", "?")
        locais = props.get("local", [])
        key = f"{rod} | {', '.join(locais) if locais else '—'}"
        resumo[key] = resumo.get(key, 0) + 1
    return resumo


# ══════════════════════════════════════
#  DASHBOARD INTERATIVO (MAPA)
# ══════════════════════════════════════
def _coords_centro_geojson(geojson_obj):
    """Calcula centro aproximado (lat, lon) do GeoJSON."""
    coords = []

    def extrair(geom):
        if not geom or not isinstance(geom, dict):
            return
        gtype = geom.get("type", "")
        c = geom.get("coordinates")
        if gtype == "Point" and isinstance(c, (list, tuple)) and len(c) >= 2:
            coords.append([float(c[1]), float(c[0])])
        elif gtype in ("LineString", "MultiPoint") and isinstance(c, (list, tuple)):
            for pt in c:
                if isinstance(pt, (list, tuple)) and len(pt) >= 2:
                    coords.append([float(pt[1]), float(pt[0])])
        elif gtype == "MultiLineString" and isinstance(c, (list, tuple)):
            for line in c:
                if isinstance(line, (list, tuple)):
                    for pt in line:
                        if isinstance(pt, (list, tuple)) and len(pt) >= 2:
                            coords.append([float(pt[1]), float(pt[0])])

    for feat in geojson_obj.get("features", []):
        extrair(feat.get("geometry"))

    if not coords:
        return (-23.55, -46.63)
    lat = sum(p[0] for p in coords) / len(coords)
    lon = sum(p[1] for p in coords) / len(coords)
    return (lat, lon)


def gerar_dashboard_artesp(caminho_geojson, titulo=None):
    """
    Gera HTML interativo com mapa Leaflet exibindo o GeoJSON (modelo SGC: header, painel KPIs, seletor de camadas).
    Retorna o caminho do arquivo .html ou None em caso de erro.
    Funciona para todos os lotes (L13, L21, L26).
    """
    if not caminho_geojson:
        return None
    path_abs = os.path.abspath(os.path.normpath(str(caminho_geojson)))
    if not os.path.isfile(path_abs):
        return None
    try:
        with open(path_abs, "r", encoding="utf-8") as f:
            geojson_obj = json.load(f)
    except Exception as e:
        import warnings
        warnings.warn("Dashboard: falha ao ler GeoJSON %s: %s" % (path_abs, e))
        return None

    lat, lon = _coords_centro_geojson(geojson_obj)
    titulo = titulo or "ARTESP — Relatório Geográfico"
    base = path_abs.replace(".geojson", "").rstrip(".json")
    html_path = f"{base}_dashboard.html"
    geojson_escaped = json.dumps(geojson_obj, ensure_ascii=False).replace("</", "<\\/")

    # KPIs dinâmicos + estatísticas por Lote e por Item (Malha Real vs Produção Total)
    features = geojson_obj.get("features") or []
    rodovias = set()
    ext_km = 0.0
    malha_fisica_set = set()
    malha_fisica_km = 0.0
    producao_por_item = {}
    for f in features:
        p = f.get("properties") or {}
        r = p.get("rodovia") or p.get("Rodovia")
        if r:
            rodovias.add(str(r).strip())
        try:
            g = f.get("geometry")
            if g and g.get("type") == "LineString":
                coords = g.get("coordinates") or []
                extensao = 0.0
                for i in range(1, len(coords)):
                    c1, c2 = coords[i - 1], coords[i]
                    extensao += ((c2[0] - c1[0]) ** 2 + (c2[1] - c1[1]) ** 2) ** 0.5 * 111  # aprox km
                ext_km += extensao
                item = (p.get("detalhamento_servico") or p.get("servico") or p.get("item") or "Outros")
                if isinstance(item, str):
                    item = item.strip() or "Outros"
                producao_por_item[item] = producao_por_item.get(item, 0) + extensao
                ki = p.get("km_inicial")
                kf = p.get("km_final")
                if ki is not None and kf is not None:
                    try:
                        key = (str(r or "").strip(), int(float(ki)), int(float(kf)))
                        if key not in malha_fisica_set:
                            malha_fisica_set.add(key)
                            malha_fisica_km += extensao
                    except (TypeError, ValueError):
                        malha_fisica_km += extensao
                else:
                    malha_fisica_km += extensao
        except Exception:
            pass
    kpi_rod = ", ".join(sorted(rodovias)[:3]) if rodovias else "Todas"
    kpi_ext = f"{ext_km:.2f} km" if ext_km > 0 else "—"
    kpi_malha_real = f"{malha_fisica_km:.2f} km" if malha_fisica_km > 0 else "—"
    # Texto "Produção por item": Roçada 400 km, Drenagem 200 km, ...
    itens_ordenados = sorted(producao_por_item.items(), key=lambda x: -x[1])
    producao_por_item_texto = ", ".join(f"{nome} {v:.1f} km" for nome, v in itens_ordenados[:8])
    if len(itens_ordenados) > 8:
        producao_por_item_texto += " …"
    producao_por_item_html = html.escape(producao_por_item_texto) if producao_por_item_texto else "—"
    data_atual = datetime.datetime.now().strftime("%d/%m/%Y")

    html_content = f"""<!DOCTYPE html>
<html lang="pt-BR">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>SGC | {titulo}</title>
  <link rel="stylesheet" href="https://unpkg.com/leaflet@1.9.4/dist/leaflet.css" />
  <script src="https://unpkg.com/leaflet@1.9.4/dist/leaflet.js"></script>
  <style>
    :root {{ --primary-bg: #0F2538; --accent-blue: #3498db; --panel-bg: rgba(255,255,255,0.95); }}
    * {{ margin: 0; padding: 0; box-sizing: border-box; }}
    body {{ font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; display: flex; flex-direction: column; height: 100vh; overflow: hidden; }}
    header {{ background: var(--primary-bg); color: white; padding: 0 20px; height: 60px; display: flex; align-items: center; justify-content: space-between; box-shadow: 0 2px 10px rgba(0,0,0,0.3); z-index: 2000; }}
    header h1 {{ font-size: 1.2rem; font-weight: 500; letter-spacing: 1px; }}
    #map {{ flex: 1; z-index: 1; }}
    .map-filters {{ position: absolute; bottom: 30px; left: 50%; transform: translateX(-50%); background: var(--panel-bg); padding: 10px 20px; border-radius: 30px; display: flex; gap: 15px; box-shadow: 0 4px 15px rgba(0,0,0,0.2); z-index: 1000; }}
    .popup-table {{ border-collapse: collapse; font-size: 12px; }}
    .popup-table th, .popup-table td {{ padding: 4px 8px; border: 1px solid #ddd; text-align: left; }}
    .popup-table th {{ background: var(--primary-bg); color: white; }}
  </style>
</head>
<body>
  <header>
    <h1>SGC | {titulo}</h1>
    <div id="date-info" style="font-size: 0.8rem; opacity: 0.8;">Atualizado em: {data_atual}</div>
  </header>
  <div id="map"></div>
  <div class="map-filters">
    <label style="font-size: 12px; font-weight: bold;">CAMADAS:</label>
    <select id="layer-select" style="border:none; outline:none; font-size:12px; background:transparent; cursor:pointer;">
      <option value="osm">OpenStreetMap (Vetor)</option>
      <option value="sat">Google Satellite (Híbrido)</option>
    </select>
  </div>
  <script>
    var geojsonData = {geojson_escaped};
    var map = L.map('map').setView([{lat}, {lon}], 12);
    var osm = L.tileLayer('https://{{s}}.tile.openstreetmap.org/{{z}}/{{x}}/{{y}}.png', {{ attribution: '&copy; OpenStreetMap contributors' }}).addTo(map);
    var sat = L.tileLayer('https://mt1.google.com/vt/lyrs=y&x={{x}}&y={{y}}&z={{z}}', {{ attribution: '&copy; Google Maps' }});
    document.getElementById('layer-select').addEventListener('change', function(e) {{
      if (e.target.value === 'sat') {{ map.addLayer(sat); map.removeLayer(osm); }}
      else {{ map.addLayer(osm); map.removeLayer(sat); }}
    }});
    function styleLine(f) {{
      var p = f.properties || {{}};
      var corFinal = p.stroke || p.color || '#3498db';
      var peso = p['stroke-width'] || 4;
      var estiloTracejado = '';
      var statusLower = (p.status || p.tipo || '').toLowerCase();
      if (statusLower === 'programado' || statusLower === 'mensal' || p.stroke === '#f1c40f') {{
        estiloTracejado = '10, 10';
      }}
      return {{ color: corFinal, weight: peso, opacity: p['stroke-opacity'] || 0.8, dashArray: estiloTracejado, lineJoin: 'round' }};
    }}
    function onEachFeature(f, layer) {{
      var p = f.properties || {{}};
      var c = '<table class="popup-table"><tr><th colspan="2">Dados de Engenharia</th></tr>' +
        '<tr><td><b>Rodovia</b></td><td>' + (p.rodovia || '—') + '</td></tr>' +
        '<tr><td><b>KM</b></td><td>' + (p.km_inicial != null ? p.km_inicial : '—') + ' - ' + (p.km_final != null ? p.km_final : '—') + '</td></tr>' +
        '<tr><td><b>Status</b></td><td>' + (p.status || 'Anual') + '</td></tr>' +
        '<tr><td><b>Serviço</b></td><td>' + (p.detalhamento_servico || p.servico || '—') + '</td></tr></table>';
      layer.bindPopup(c);
      layer.on('mouseover', function() {{ this.setStyle({{ weight: 7, color: '#FFFFFF', opacity: 1 }}); }});
      layer.on('mouseout', function() {{ this.setStyle(styleLine(f)); }});
    }}
    var geoLayer = L.geoJSON(geojsonData, {{
      style: styleLine,
      onEachFeature: onEachFeature,
      pointToLayer: function(f, latlng) {{ return L.circleMarker(latlng, {{ radius: 6, fillColor: '#D32F2F', color: '#fff', weight: 2 }}); }}
    }}).addTo(map);
    if (geoLayer.getBounds && geoLayer.getBounds().isValid()) map.fitBounds(geoLayer.getBounds(), {{ padding: [30, 30] }});
  </script>
</body>
</html>"""

    try:
        with open(html_path, "w", encoding="utf-8") as f:
            f.write(html_content)
        abs_path = os.path.abspath(html_path)
        if os.path.isfile(html_path):
            print("  [OK] Dashboard gravado: %s" % abs_path)
        else:
            print("  [AVISO] Dashboard: ficheiro não encontrado após write: %s" % abs_path)
        return html_path
    except Exception as e:
        print("  [ERRO] Dashboard: falha ao gravar HTML %s: %s" % (os.path.abspath(html_path), e))
        import warnings
        warnings.warn("Dashboard: falha ao gravar HTML %s: %s" % (html_path, e))
        return None


# ══════════════════════════════════════════════════════
#  CONSTRUIR TEXTO DO RELATÓRIO LOG
# ══════════════════════════════════════════════════════
def _construir_texto_relatorio_log(versao, lote_sigla, modalidade_rotulo, n_linhas_df, n_features, n_pendencias,
                                   resumo_dict, sha_schema, sha_geojson, sha_excel, validacao_resultado,
                                   correcao_eixo, n_marcadores=0, sentidos_invertidos=None):
    linhas = []
    linhas.append("=" * 70)
    linhas.append(f"RELATÓRIO DE GERAÇÃO - GEOJSON DO LOTE {lote_sigla}")
    linhas.append("=" * 70)
    linhas.append("")
    linhas.append(f"Data/Hora: {datetime.datetime.now().strftime('%d/%m/%Y %H:%M:%S')}")
    linhas.append(f"Lote: {lote_sigla}")
    linhas.append(f"Modalidade: {modalidade_rotulo}")
    linhas.append("Versão Schema: R0")
    linhas.append("")
    if lote_sigla not in LOTES_SCHEMA_VALIDOS:
        linhas.append("[!] AVISO IMPORTANTE:")
        linhas.append(f"   {lote_sigla} NÃO ESTÁ NA LISTA DE LOTES DO SCHEMA OFICIAL!")
        linhas.append(f"   Lotes válidos: {', '.join(LOTES_SCHEMA_VALIDOS)}")
        linhas.append("")
    if sentidos_invertidos:
        linhas.append("-" * 70)
        linhas.append("SENTIDOS INVERTIDOS DETECTADOS (corrigidos automaticamente):")
        linhas.append("-" * 70)
        for rod, sentido in sentidos_invertidos:
            linhas.append(f"  {rod} — {sentido}")
        linhas.append("")
    linhas.append("-" * 70)
    linhas.append("ESTATÍSTICAS:")
    linhas.append("-" * 70)
    linhas.append(f"Linhas Excel: {n_linhas_df}")
    linhas.append(f"Features (trechos) geradas: {n_features}")
    if n_marcadores:
        linhas.append(f"Marcadores (alfinetes): {n_marcadores}")
    razao = n_features / n_linhas_df if n_linhas_df > 0 else 0
    linhas.append(f"Razão: {razao:.2f} feature/linha")
    linhas.append(f"Pendências: {n_pendencias}")
    linhas.append(f"Correção de eixo: {'ATIVADA' if correcao_eixo else 'DESATIVADA'}")
    linhas.append("")
    linhas.append("-" * 70)
    linhas.append("FEATURES POR RODOVIA/SENTIDO:")
    linhas.append("-" * 70)
    for k, v in (resumo_dict or {}).items():
        linhas.append(f"  {k:30} → {v:4} features")
    linhas.append("")
    linhas.append("-" * 70)
    linhas.append("INTEGRIDADE DOS ARQUIVOS (SHA256):")
    linhas.append("-" * 70)
    linhas.append(f"Schema JSON:  {sha_schema}")
    linhas.append(f"GeoJSON:      {sha_geojson}")
    linhas.append(f"Excel:        {sha_excel}")
    linhas.append("")
    linhas.append("-" * 70)
    linhas.append("VALIDAÇÃO JSON SCHEMA:")
    linhas.append("-" * 70)
    status = "APROVADO" if (validacao_resultado and validacao_resultado.get('ok')) else "NEGADO"
    linhas.append(f"STATUS: {status}")
    linhas.append("")
    linhas.append(validacao_resultado.get('msg', 'N/A') if validacao_resultado else 'N/A')
    linhas.append("")
    linhas.append("=" * 70)
    linhas.append("FIM DO RELATÓRIO")
    linhas.append("=" * 70)
    return "\n".join(linhas)


# ══════════════════════════════════════
#  GERAR PDF RELATÓRIO
# ══════════════════════════════════════
def gerar_pdf_relatorio(features, lote_info, caminho_pdf, resumo_dict, cabecalho_5_linhas=None,
                        validacao_resultado=None, relatorio_log_data=None):
    if not REPORTLAB_OK:
        print("[AVISO] ReportLab não disponível - PDF não gerado")
        return False

    if not relatorio_log_data:
        print("[AVISO] PDF: sem dados do relatório (relatorio_log_data)")
        return False

    try:
        doc = SimpleDocTemplate(
            caminho_pdf,
            pagesize=A4,
            rightMargin=1.5 * cm,
            leftMargin=1.5 * cm,
            topMargin=1.5 * cm,
            bottomMargin=1.5 * cm
        )

        story = []
        styles = getSampleStyleSheet()

        n_pend = relatorio_log_data.get("n_pendencias", 0)
        versao_key = relatorio_log_data.get("versao_key", "r0")
        n_linhas = relatorio_log_data.get("n_linhas_df", 0)

        style_titulo = ParagraphStyle(
            "ResumoTitulo",
            parent=styles["Normal"],
            fontName="Helvetica-Bold",
            fontSize=12,
            spaceAfter=6,
        )
        style_resumo = ParagraphStyle(
            "ResumoTexto",
            parent=styles["Normal"],
            fontName="Helvetica",
            fontSize=10,
            spaceAfter=4,
        )
        style_alerta = ParagraphStyle(
            "ResumoAlerta",
            parent=styles["Normal"],
            fontName="Helvetica",
            fontSize=10,
            textColor=colors.HexColor("#CC0000"),
            spaceAfter=4,
        )
        style_ok = ParagraphStyle(
            "ResumoOk",
            parent=styles["Normal"],
            fontName="Helvetica",
            fontSize=10,
            textColor=colors.HexColor("#008800"),
            spaceAfter=4,
        )

        story.append(Paragraph(f"Resumo do Processamento — Versão {versao_key}", style_titulo))
        story.append(Paragraph(f"Total de Registros Processados: {n_linhas}", style_resumo))
        if n_pend > 0:
            story.append(Paragraph(
                f"Pendências Encontradas: {n_pend} (Ver anexo PENDENCIAS.csv)",
                style_alerta,
            ))
        else:
            story.append(Paragraph(
                "Status: 100% de Conformidade Geográfica.",
                style_ok,
            ))
        story.append(Spacer(1, 12))

        texto_log = _construir_texto_relatorio_log(
            relatorio_log_data.get("versao", VERSAO),
            relatorio_log_data.get("lote_sigla", lote_info.get("sigla", "")),
            relatorio_log_data.get("modalidade_rotulo", ""),
            relatorio_log_data.get("n_linhas_df", 0),
            relatorio_log_data.get("n_features", len(features)),
            relatorio_log_data.get("n_pendencias", 0),
            relatorio_log_data.get("resumo", resumo_dict),
            relatorio_log_data.get("sha_schema", ""),
            relatorio_log_data.get("sha_geojson", ""),
            relatorio_log_data.get("sha_excel", ""),
            relatorio_log_data.get("validacao_resultado", validacao_resultado),
            relatorio_log_data.get("correcao_eixo", False),
            n_marcadores=relatorio_log_data.get("n_marcadores", 0),
            sentidos_invertidos=relatorio_log_data.get("sentidos_invertidos_detectados") or (),
        )
        style_mono = ParagraphStyle(
            'RelatorioLog',
            parent=styles['Normal'],
            fontName='Courier',
            fontSize=8,
            leading=9,
            leftIndent=0,
            rightIndent=0,
        )
        for linha in texto_log.split("\n"):
            esc = (linha or " ").replace("&", "&").replace("<", "<").replace(">", ">")
            story.append(Paragraph(esc, style_mono))

        doc.build(story)
        print(f"[OK] PDF (retrato - log): {caminho_pdf}")
        return True

    except Exception as e:
        print(f"[ERRO] PDF: {e}")
        import traceback
        traceback.print_exc()
        return False


# ══════════════════════════════════════
#  LICENÇA
# ══════════════════════════════════════
def _id_mquina():
    """Gera ID único da máquina."""
    import uuid as _uuid
    try:
        mac = _uuid.getnode()
        return hashlib.sha256(f"{mac}{platform.node()}".encode()).hexdigest()[:32]
    except Exception:
        return "UNKNOWN"


def _exige_verificacao_licenca():
    return getattr(sys, "frozen", False)


def _verificar_licenca_stub():
    """Stub para compatibilidade (arquivo truncado)."""
    return True