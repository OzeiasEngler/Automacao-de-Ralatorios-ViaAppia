"""
nc_artesp/modulos/analisar_pdf_nc.py
────────────────────────────────────────────────────────────────────────────
Analisa o PDF de Constatação de Rotina Artesp:
  1. Extrai os apontamentos (NCs) do texto do PDF sem alterar nada.
  2. Ordena por rodovia → sentido → km e detecta saltos de KM.
  3. Alerta sobre trechos sem apontamento onde panelas (buraco/recalque)
     podem estar não registradas — prazo legal: 24 h.
  4. Gera relatório PDF com:
       • Capa com resumo e estatísticas
       • Alertas de salto de KM
       • NCs emergenciais (prazo ≤ 24 h)
       • NCs agrupadas por tipo, em sequência de KM, dados originais intactos

Dependências: PyMuPDF (fitz), reportlab — já presentes no requirements.txt.
Desenvolvedor: Ozeias Engler
"""

from __future__ import annotations

import io
import logging
import os
import re
import unicodedata
from dataclasses import dataclass, field
from datetime import datetime, date
from pathlib import Path
from typing import Optional

logger = logging.getLogger(__name__)

# ── PyMuPDF ───────────────────────────────────────────────────────────────────
try:
    import fitz
    FITZ_OK = True
except ImportError:
    FITZ_OK = False

# ── ReportLab ────────────────────────────────────────────────────────────────
try:
    from reportlab.lib import colors
    from reportlab.lib.pagesizes import A4
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.lib.units import mm
    from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_RIGHT
    from reportlab.platypus import (
        SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle,
        HRFlowable, PageBreak, KeepTogether,
    )
    REPORTLAB_OK = True
except ImportError:
    REPORTLAB_OK = False

try:
    import openpyxl
    OPENPYXL_OK = True
except ImportError:
    openpyxl = None
    OPENPYXL_OK = False

# ─────────────────────────────────────────────────────────────────────────────
# CONFIGURAÇÕES
# ─────────────────────────────────────────────────────────────────────────────

LIMIAR_GAP_KM   = 2.0    # gap entre NCs consecutivas (km) para gerar alerta
PRAZO_EMERG_MAX = 1      # prazo ≤ 1 dia = emergencial (24 h)

# Carrega o mapa EAF do config (pode ser sobreposto via parâmetro)
try:
    from nc_artesp.config import MAPA_EAF as _MAPA_EAF_PADRAO
except ImportError:
    try:
        from config import MAPA_EAF as _MAPA_EAF_PADRAO
    except ImportError:
        _MAPA_EAF_PADRAO = []

# Palavras-chave que identificam possível panela/buraco no tipo de NC
PALAVRAS_PANELA = [
    "buraco", "panela", "remendo", "afundamento", "recalque",
    "ruptura", "deformação", "deformacao", "trinca", "fissura",
    "escorregamento", "deslizamento", "depressão", "depressao",
    "irregular", "desgaste", "trilha",
]

# Atividades cujo prazo legal ARTESP é 24 h (buraco/panela)
ATIVIDADES_24H = [
    "buraco", "panela", "remendo profundo", "recapeamento de buraco",
]


# ─────────────────────────────────────────────────────────────────────────────
# MODELO DE DADOS
# ─────────────────────────────────────────────────────────────────────────────

@dataclass
class NcItem:
    codigo: str           = ""
    data_con: str         = ""   # DD/MM/AAAA
    horario_fiscalizacao: str = ""  # HH:MM:SS (col 5 do relatório)
    km_ini_str: str       = ""   # "50 + 950"
    km_fim_str: str       = ""
    km_ini: float         = 0.0  # valor decimal (50.950)
    km_fim: float         = 0.0
    sentido: str          = ""
    atividade: str        = ""   # descrição completa (col 17)
    tipo_atividade: str   = ""   # ex.: Faixa de Domínio (col 15)
    grupo_atividade: str  = ""   # ex.: Limpeza (col 16)
    observacao: str       = ""
    rodovia: str          = ""   # "SP 075"
    rodovia_nome: str     = ""
    lote: str             = ""
    concessionaria: str   = ""
    prazo_str: str        = ""   # DD/MM/AAAA
    prazo_dias: Optional[int] = None
    emergencial: bool     = False
    tipo_panela: bool     = False  # atividade sugere panela/buraco
    grupo: int            = 0    # grupo EAF (col V template)
    empresa: str          = ""   # nome da empresa fiscalizadora (col 20 EAF)
    nome_fiscal: str      = ""   # responsável técnico (col 21)
    origem_ma: bool       = False


@dataclass
class GapAlerta:
    """Salto de KM dentro de um mesmo grupo/rodovia/sentido."""
    grupo: int
    empresa: str
    rodovia: str
    sentido: str
    km_antes: float
    km_depois: float
    gap_km: float
    nc_antes: str    # código fiscalização
    nc_depois: str


@dataclass
class CodigoGapAlerta:
    """Salto na numeração sequencial do Código Fiscalização dentro de um grupo.
    Indica apontamentos que foram gerados mas não foram entregues à concessionária
    — potencialmente buracos/panelas na pista (prazo 24 h)."""
    grupo: int
    empresa: str
    codigo_antes: str        # último código entregue
    codigo_depois: str       # próximo código entregue
    n_faltantes: int         # quantos estão faltando entre os dois
    codigos_faltantes: list  # lista dos códigos ausentes (até 10 por alerta)


# ─────────────────────────────────────────────────────────────────────────────
# PARSER DO PDF
# ─────────────────────────────────────────────────────────────────────────────

def _km_para_float(s: str) -> float:
    """'50 + 950' → 50.950 ; '67 + 000' → 67.0"""
    m = re.match(r'(\d+)\s*\+\s*(\d+)', s.strip())
    if m:
        return int(m.group(1)) + int(m.group(2)) / 1000.0
    try:
        return float(s.replace(",", "."))
    except Exception:
        return 0.0


def _parse_data(s: str) -> Optional[datetime]:
    for fmt in ("%d/%m/%Y", "%d/%m/%y"):
        try:
            return datetime.strptime(s.strip(), fmt)
        except ValueError:
            pass
    return None


def _prazo_dias(data_con: str, prazo: str) -> Optional[int]:
    d1 = _parse_data(data_con)
    d2 = _parse_data(prazo)
    if d1 and d2:
        return max(0, (d2 - d1).days)
    return None


def _is_panela(atividade: str) -> bool:
    a = atividade.lower()
    return any(p in a for p in PALAVRAS_PANELA)


# Grupo = resumo da atividade; Tipo = descrição do tipo. (palavras na atividade, grupo, tipo)
_MAPA_ATIVIDADE_GRUPO_TIPO: list[tuple[list[str], str, str]] = [
    (["erosao", "erosão", "bueiro", "ruptura de bueiro"], "Erosão", "Recuperação de erosão"),
    (["defensa", "defensas", "acidentada", "reposição de defensa"], "Defensas", "Reposição de defensa acidentada"),
    (["pavimento", "buraco", "panela", "recapeamento", "depressao", "depressão"], "Pavimento", "Pavimento"),
    (["cerca", "cercas", "alambrado", "telamento", "vedos"], "Segurança", "Segurança"),
    (["limpeza", "lixo", "faixa de dominio", "faixa de domínio"], "Limpeza", "Faixa de Domínio"),
    (["drenagem", "dreno", "sarjeta", "galeria"], "Drenagem", "Drenagem"),
    (["sinalizacao", "sinalização", "conifica", "conificação", "fita zebrada"], "Sinalização", "Sinalização"),
]


def _inferir_grupo_tipo_da_atividade(atividade: str) -> tuple[str, str]:
    """Infere grupo (resumo) e tipo de atividade a partir do texto da atividade quando não preenchidos."""
    if not (atividade or "").strip():
        return ("", "")
    a = (atividade or "").strip().lower()
    for keywords, grupo, tipo in _MAPA_ATIVIDADE_GRUPO_TIPO:
        if any(kw in a for kw in keywords):
            return (grupo, tipo)
    return ("", "")


# Mínimo de caracteres por página para considerar que há texto nativo (senão tenta OCR)
_EXTRAIR_TEXTO_MIN_LEN = 40


def _extrair_texto_pdf(pdf_bytes: bytes) -> str:
    """Extrai todo o texto do PDF em uma string única. Inclui versão por blocos para tabelas.
    Para PDFs escaneados (sem camada de texto), usa OCR (pytesseract) se disponível."""
    if not FITZ_OK:
        raise ImportError("PyMuPDF não instalado: pip install pymupdf")
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    partes = []
    partes_blocks = []
    for pag in doc:
        txt = pag.get_text("text") or ""
        if len(txt.strip()) < _EXTRAIR_TEXTO_MIN_LEN:
            try:
                from nc_artesp.pdf_ocr import texto_de_pagina_ocr
                ocr = texto_de_pagina_ocr(pag, dpi=200)
                if ocr:
                    txt = ocr
            except Exception as e:
                logger.debug("OCR página %s: %s", pag.number + 1, e)
        partes.append(txt)
        try:
            blocs = pag.get_text("blocks")
            if blocs:
                # Ordenar por y depois x (ordem de leitura) e juntar texto de cada bloco
                blocs.sort(key=lambda b: (round(b[1], 0), round(b[0], 0)))
                blocos_str = "\n".join((b[4] or "").strip() for b in blocs if (b[4] or "").strip())
                if len(blocos_str.strip()) < _EXTRAIR_TEXTO_MIN_LEN and txt.strip():
                    blocos_str = txt
                partes_blocks.append(blocos_str)
            elif txt.strip():
                partes_blocks.append(txt)
        except Exception:
            if txt.strip():
                partes_blocks.append(txt)
    doc.close()
    texto = "\n".join(partes)
    if partes_blocks:
        texto_blocks = "\n".join(partes_blocks)
        if texto_blocks and texto_blocks != texto:
            texto = texto + "\n" + texto_blocks
    return texto


def _sentido_para_texto(s: str) -> str:
    """Converte letra (L/O/N/S/I/E) para nome completo no relatório de análise."""
    s = (s or "").strip().upper()
    if not s:
        return ""
    letra = s[0] if s else ""
    if letra == "0":
        letra = "O"
    mapa = {"L": "Leste", "O": "Oeste", "N": "Norte", "S": "Sul", "I": "Interna", "E": "Externa"}
    return mapa.get(letra, s)


def _parse_nc_block(block: str) -> Optional[NcItem]:
    """
    Analisa um bloco de texto correspondente a uma NC.

    Estrutura observada no PDF ARTESP (com variações de layout/OCR):
      {data_con}   Constatação -
      Código Fiscalização: Lote: {codigo} Concessionária: {conc}
      {km_ini}   Km+m - Inicial: {km_fim}   Km+m - Final: {sentido}   Sentido:
      Data Limite para Reparo -
      Atividade: {atividade}
      [{observacao}]   Observação: / Obs.:
      Rodovia (SP): {rodovia}
    """
    nc = NcItem()
    lines = [ln.strip() for ln in block.splitlines() if ln.strip()]
    # Uma única linha com espaços (para regex que cruzam quebra de linha)
    texto_flat = " ".join(lines)

    # ── data constatação (primeira linha com padrão de data)
    for ln in lines[:5]:
        m = re.match(r'^(\d{2}/\d{2}/\d{4})', ln)
        if m:
            nc.data_con = m.group(1)
            break
    if not nc.data_con:
        m = re.search(r'Data\s+(?:Constata[cç][aã]o\s*[:\-]?\s*)?(\d{2}/\d{2}/\d{4})', texto_flat, re.IGNORECASE)
        if m:
            nc.data_con = m.group(1)
    if not nc.data_con:
        for i, ln in enumerate(lines[:8]):
            if re.match(r'^Data\s*$', ln, re.IGNORECASE) or re.match(r'^Data\s+Constata[cç][aã]o\s*$', ln, re.IGNORECASE):
                if i + 1 < len(lines) and re.match(r'^\d{2}/\d{2}/\d{4}', lines[i + 1]):
                    nc.data_con = lines[i + 1].strip()[:10]
                    break

    # ── código fiscalização + concessionária (vários formatos)
    for ln in lines:
        m = re.search(
            r'C[oó]digo\s+Fiscaliza[cç][aã]o:\s*Lote:\s*(\S+)\s+Concession[aá]ria:\s*(.+)',
            ln, re.IGNORECASE
        )
        if m:
            nc.codigo = m.group(1).strip()
            nc.concessionaria = m.group(2).strip()
            break
    if not nc.codigo:
        for ln in lines:
            m = re.search(r'Cod\.?\s*Fiscal\.?\s*:?\s*(\d+)', ln, re.IGNORECASE)
            if m:
                nc.codigo = m.group(1).strip()
                break
    if not nc.codigo:
        m = re.search(r'Lote:\s*(\d+)', texto_flat, re.IGNORECASE)
        if m:
            nc.codigo = m.group(1).strip()
        if not nc.codigo:
            m = re.search(r'C[oó]digo\s+Fiscal[.:]?\s*(\d+)', texto_flat, re.IGNORECASE)
            if m:
                nc.codigo = m.group(1).strip()
    if not nc.codigo:
        m = re.search(r'C[oó]digo(?!\s+da)(?!\s+Fiscaliza)\s*:?\s*(\d{4,})', texto_flat, re.IGNORECASE)
        if m:
            nc.codigo = m.group(1).strip()
    if not nc.codigo:
        m = re.search(r'(\d{4,})\s+C[oó]digo(?!\s+da)(?!\s+Fiscaliza)', texto_flat, re.IGNORECASE)
        if m:
            nc.codigo = m.group(1).strip()
    if not nc.codigo:
        for i, ln in enumerate(lines):
            if re.match(r'^C[oó]digo\s*:?\s*$', ln, re.IGNORECASE):
                if i + 1 < len(lines) and re.match(r'^\d{4,}\s*$', lines[i + 1].strip()):
                    nc.codigo = lines[i + 1].strip()
                    break
    if not nc.codigo:
        for i, ln in enumerate(lines):
            if re.match(r'^Codigo\s*:?\s*$', ln, re.IGNORECASE):
                if i + 1 < len(lines) and re.match(r'^\d{4,}\s*$', lines[i + 1].strip()):
                    nc.codigo = lines[i + 1].strip()
                    break
    if not nc.codigo and re.search(r'\bC[oó]digo\b', texto_flat, re.IGNORECASE):
        for i, ln in enumerate(lines[:12]):
            if not re.match(r'^\d{4,}\s*$', ln.strip()) or ln.strip().upper().startswith("LOTE"):
                continue
            adj = (lines[i - 1] if i > 0 else "") + " " + (lines[i + 1] if i + 1 < len(lines) else "")
            if re.search(r'C[oó]digo', adj, re.IGNORECASE):
                nc.codigo = ln.strip()
                break
    # Layout emergencial: "Código" numa linha, valor (5+ dígitos) mais abaixo, antes de "Lote:" ou "Data da Constatação"
    if not nc.codigo and re.search(r'\bC[oó]digo\b', texto_flat, re.IGNORECASE):
        idx_codigo = next((i for i, ln in enumerate(lines) if re.match(r'^C[oó]digo\s*$', ln, re.IGNORECASE)), -1)
        if idx_codigo >= 0:
            for ln in lines[idx_codigo + 1: min(idx_codigo + 15, len(lines))]:
                if re.match(r'^Lote\s*:?\s*', ln, re.IGNORECASE) or re.match(r'^Data\s+da\s+Constata', ln, re.IGNORECASE):
                    break
                if re.match(r'^\d{5,}\s*$', ln.strip()):
                    nc.codigo = ln.strip()
                    break

    # ── km inicial / km final / sentido (uma linha ou texto contínuo)
    for ln in lines:
        m = re.search(
            r'(\d+\s*\+\s*\d+)\s+Km\+m\s*-\s*Inicial:\s*(\d+\s*\+\s*\d+)\s+Km\+m\s*-\s*Final:\s*(.+?)\s+Sentido:',
            ln, re.IGNORECASE
        )
        if m:
            nc.km_ini_str = m.group(1).strip()
            nc.km_fim_str = m.group(2).strip()
            nc.sentido = _sentido_para_texto(m.group(3).strip())
            nc.km_ini = _km_para_float(nc.km_ini_str)
            nc.km_fim = _km_para_float(nc.km_fim_str)
            break
    if not nc.km_ini_str:
        m = re.search(
            r'(\d+\s*\+\s*\d+)\s+Km\+m\s*-\s*Inicial:\s*(\d+\s*\+\s*\d+)\s+Km\+m\s*-\s*Final:\s*(.+?)\s+Sentido:',
            texto_flat, re.IGNORECASE
        )
        if m:
            nc.km_ini_str = m.group(1).strip()
            nc.km_fim_str = m.group(2).strip()
            nc.sentido = _sentido_para_texto(m.group(3).strip())
            nc.km_ini = _km_para_float(nc.km_ini_str)
            nc.km_fim = _km_para_float(nc.km_fim_str)
    if not nc.km_ini_str:
        # Alternativa: Km (sem +m), Inicial/Final em qualquer ordem
        m = re.search(
            r'(\d+\s*\+\s*\d+)\s+Km\s*[-\s]*Inicial\s*:?\s*(\d+\s*\+\s*\d+)\s+Km\s*[-\s]*Final\s*:?\s*([LONSIE0])\s*Sentido',
            texto_flat, re.IGNORECASE
        )
        if m:
            nc.km_ini_str = m.group(1).strip()
            nc.km_fim_str = m.group(2).strip()
            nc.sentido = _sentido_para_texto(m.group(3).strip())
            nc.km_ini = _km_para_float(nc.km_ini_str)
            nc.km_fim = _km_para_float(nc.km_fim_str)
    if not nc.km_ini_str:
        # Padrão mais solto: dois "km+m" e depois sentido
        m = re.search(
            r'(\d+\s*\+\s*\d+)\s+.*?Inicial\s*:?\s*(\d+\s*\+\s*\d+)\s+.*?Final\s*:?\s*([LONSIE0]?)\s*Sentido',
            texto_flat, re.IGNORECASE | re.DOTALL
        )
        if m:
            nc.km_ini_str = m.group(1).strip()
            nc.km_fim_str = m.group(2).strip()
            nc.sentido = _sentido_para_texto((m.group(3) or "").strip())
            nc.km_ini = _km_para_float(nc.km_ini_str)
            nc.km_fim = _km_para_float(nc.km_fim_str)
    if not nc.km_ini_str and re.search(r'Km|Inicial|Final|Sentido', texto_flat, re.I):
        # Último recurso: dois "número + número" (km ini e km fim) no bloco
        m = re.findall(r'(\d+\s*\+\s*\d+)', texto_flat)
        if len(m) >= 2:
            nc.km_ini_str = m[0].strip()
            nc.km_fim_str = m[1].strip()
            nc.km_ini = _km_para_float(nc.km_ini_str)
            nc.km_fim = _km_para_float(nc.km_fim_str)
        elif len(m) == 1:
            nc.km_ini_str = nc.km_fim_str = m[0].strip()
            nc.km_ini = nc.km_fim = _km_para_float(nc.km_ini_str)
    if not nc.km_ini_str and re.search(r'Km\+m', texto_flat, re.I):
        for i, ln in enumerate(lines):
            if re.search(r'Km\+m', ln, re.IGNORECASE):
                for j in range(i + 1, min(i + 5, len(lines))):
                    km_m = re.search(r'(\d+\s*\+\s*\d+)', lines[j])
                    if km_m:
                        nc.km_ini_str = nc.km_fim_str = km_m.group(1).strip()
                        nc.km_ini = nc.km_fim = _km_para_float(nc.km_ini_str)
                        break
                break
    if not nc.km_ini_str or not nc.km_fim_str:
        for i, ln in enumerate(lines):
            if re.match(r'^Inicial\s*:?\s*$', ln, re.IGNORECASE) and i + 1 < len(lines):
                km = re.search(r'(\d+\s*\+\s*\d+)', lines[i + 1])
                if km and not nc.km_ini_str:
                    nc.km_ini_str = km.group(1).strip()
                    nc.km_ini = _km_para_float(nc.km_ini_str)
            if re.match(r'^Final\s*:?\s*$', ln, re.IGNORECASE) and i + 1 < len(lines):
                km = re.search(r'(\d+\s*\+\s*\d+)', lines[i + 1])
                if km and not nc.km_fim_str:
                    nc.km_fim_str = km.group(1).strip()
                    nc.km_fim = _km_para_float(nc.km_fim_str)
    if not nc.sentido and nc.km_ini_str:
        m = re.search(r'Sentido\s*:?\s*([LONSIE0])\b', texto_flat, re.IGNORECASE)
        if m:
            nc.sentido = _sentido_para_texto(m.group(1).strip())
        if not nc.sentido:
            m = re.search(r'\b([LONSIE0])\s+Sentido\b', texto_flat, re.IGNORECASE)
            if m:
                nc.sentido = _sentido_para_texto(m.group(1).strip())
    if not nc.sentido:
        for i, ln in enumerate(lines):
            if re.match(r'^Sentido\s*:?\s*$', ln, re.IGNORECASE) and i + 1 < len(lines):
                next_ln = lines[i + 1].strip()
                s = re.search(r'\b([LONSIE0])\b', next_ln, re.IGNORECASE)
                if s:
                    nc.sentido = _sentido_para_texto(s.group(1).strip())
                    break
                if next_ln.lower() in ("sul", "norte", "leste", "oeste", "interna", "externa"):
                    nc.sentido = next_ln
                    break

    # ── lote (formato emergencial "Lote: 13" ou após "Data Limite para Reparo"; não confundir com "Lote: 896643" do código)
    m = re.search(r'Lote\s*:?\s*(\d+)', texto_flat, re.IGNORECASE)
    if m:
        val = m.group(1).strip()
        if val != (nc.codigo or "").strip() or len(val) <= 3:
            nc.lote = val
    if not nc.lote:
        for i, ln in enumerate(lines):
            if re.match(r'^Lote\s*:?\s*$', ln, re.IGNORECASE) and i + 1 < len(lines):
                next_val = lines[i + 1].strip()
                if re.match(r'^\d+$', next_val) and (next_val != (nc.codigo or "").strip() or len(next_val) <= 3):
                    nc.lote = next_val
                    break
    if not nc.lote:
        limite_idx = next(
            (i for i, ln in enumerate(lines) if "Limite para Reparo" in ln), -1
        )
        if limite_idx >= 0:
            for ln in lines[limite_idx + 1: limite_idx + 4]:
                if re.match(r'^\d+$', ln.strip()):
                    nc.lote = ln.strip()
                    break

    # ── atividade (padrões normais primeiro; "Evento:" só como fallback para PDF emergencial)
    for ln in lines:
        m = re.match(r'Atividade\s*:?\s*(.+?)(?=\s*Grupo\s|\s*Tipo\s|$)', ln, re.IGNORECASE | re.DOTALL)
        if m and m.group(1).strip():
            nc.atividade = m.group(1).strip()
            break
    if not nc.atividade:
        for ln in lines:
            m = re.match(r'Atividade:\s*(.+)', ln, re.IGNORECASE)
            if m:
                nc.atividade = m.group(1).strip()
                break
    if not nc.atividade:
        for i, ln in enumerate(lines):
            if re.match(r'^Atividade\s*:?\s*$', ln, re.IGNORECASE):
                if i + 1 < len(lines) and lines[i + 1].strip():
                    nc.atividade = lines[i + 1].strip()[:200]
                    break
    if not nc.atividade:
        for ln in lines:
            m = re.search(r'Evento\s*:?\s*(.+)', ln, re.IGNORECASE)
            if m and m.group(1).strip():
                nc.atividade = m.group(1).strip()[:200]
                break
    if not nc.atividade:
        for i, ln in enumerate(lines):
            if re.match(r'^Evento\s*:?\s*$', ln, re.IGNORECASE):
                if i + 1 < len(lines) and lines[i + 1].strip():
                    nc.atividade = lines[i + 1].strip()[:200]
                    break

    # ── horário da fiscalização (col E): consta nos PDFs como "Horário da(s) constatação(ões)" ou "Horário da Fiscalização"
    _padrao_hora = r'(\d{1,2}:\d{2}(?::\d{2})?)'
    _padrao_hora_h = r'(\d{1,2})h\s*(\d{2})'
    m = re.search(r'Hor[áa]rio\s*(?:da\s*)?Fiscaliza[cç][aã]o\s*:?\s*' + _padrao_hora, texto_flat, re.IGNORECASE)
    if m:
        nc.horario_fiscalizacao = m.group(1).strip()
    if not nc.horario_fiscalizacao:
        m = re.search(r'Hor[áa]rio\s*(?:das?\s*)?Constata[cç][oõ]es?\s*:?\s*' + _padrao_hora, texto_flat, re.IGNORECASE)
        if m:
            nc.horario_fiscalizacao = m.group(1).strip()
    if not nc.horario_fiscalizacao:
        m = re.search(r'Hora\s*:?\s*' + _padrao_hora, texto_flat, re.IGNORECASE)
        if m:
            nc.horario_fiscalizacao = m.group(1).strip()
    if not nc.horario_fiscalizacao:
        m = re.search(r'Constata[cç][aã]o\s*[:\-]?\s*' + _padrao_hora, texto_flat, re.IGNORECASE)
        if m:
            nc.horario_fiscalizacao = m.group(1).strip()
        if not nc.horario_fiscalizacao:
            m = re.search(_padrao_hora + r'\s*[-\s]*Constata[cç][aã]o', texto_flat, re.IGNORECASE)
            if m:
                nc.horario_fiscalizacao = m.group(1).strip()
    if not nc.horario_fiscalizacao:
        m = re.search(_padrao_hora, texto_flat)
        if m:
            nc.horario_fiscalizacao = m.group(1)
    if not nc.horario_fiscalizacao:
        m = re.search(r'(\d{1,2})h\s*(\d{2})', texto_flat)
        if m:
            nc.horario_fiscalizacao = "{}:{}".format(m.group(1), m.group(2))
    for i, ln in enumerate(lines):
        if re.search(r'Hor[áa]rio\s*(?:da\s*)?(?:Fiscaliza[cç][aã]o|Constata[cç][oõ]es?)|Hora\b|Constata[cç][aã]o\s*-\s*$', ln, re.IGNORECASE):
            m = re.search(_padrao_hora, ln)
            if m:
                nc.horario_fiscalizacao = m.group(1).strip()
                break
            if i + 1 < len(lines):
                m = re.search(r'(\d{1,2}:\d{2}(?::\d{2})?|\d{1,2}h\s*\d{2})', lines[i + 1])
                if m:
                    h = m.group(1).replace("h", ":").replace(" ", "")
                    nc.horario_fiscalizacao = h
                    break
            break
    for i, ln in enumerate(lines):
        if not nc.horario_fiscalizacao and re.search(r'\bConstata[cç][aã]o\b', ln, re.IGNORECASE):
            m = re.search(_padrao_hora, ln)
            if m:
                nc.horario_fiscalizacao = m.group(1).strip()
                break
            if i + 1 < len(lines):
                m = re.search(_padrao_hora, lines[i + 1])
                if m:
                    nc.horario_fiscalizacao = m.group(1).strip()
                    break

    # ── tipo de atividade (col O) / grupo de atividade (col P)
    # PDF pode vir: "Atividade ... Grupo Pavimento Tipo Faixa de domínio" ou "Grupo: Limpeza Tipo: Segurança"
    # Valores comuns: Tipo = Faixa de Domínio, Pavimento, Segurança; Grupo = Limpeza, Pavimento, Vedos Cercas Alambrados e Telamento
    def _limpar_valor(s: str, max_len: int) -> str:
        s = (s or "").strip()
        for prefix in ("Atividade", "Grupo", "Tipo", "Observação", "Obs"):
            m = re.match(re.escape(prefix) + r"\s*[,:]?\s*", s, re.IGNORECASE)
            if m:
                s = s[m.end() :].strip()
                break
        return s[:max_len] if s else ""

    # texto_flat: padrões "Tipo de Atividade: X" / "Grupo de Atividade: X" ou "Tipo X" / "Grupo X"
    # Parar antes de Grupo, Tipo, Atividade, Data, Observação (para não capturar o próximo campo)
    m = re.search(
        r'Tipo\s*(?:de\s*)?[Aa]tividade\s*[,:]?\s*([^\n|;]+?)(?=\s*Grupo\s|\s*Atividade\s*[,:]|\s*Data\s|Observa[cç][aã]o|\s*$)',
        texto_flat, re.IGNORECASE
    )
    if m:
        nc.tipo_atividade = _limpar_valor(m.group(1), 200)
    m = re.search(
        r'Grupo\s*(?:de\s*)?[Aa]tividade\s*[,:]?\s*([^\n|;]+?)(?=\s*Tipo\s|\s*Atividade\s*[,:]|\s*Data\s|Observa[cç][aã]o|\s*$)',
        texto_flat, re.IGNORECASE
    )
    if m:
        nc.grupo_atividade = _limpar_valor(m.group(1), 100)
    # "Tipo" / "Grupo" sozinhos: "Tipo Faixa de dominio", "Grupo Limpeza", "Tipo, Segurança", "Grupo, Vedos, Cercas..."
    if not nc.tipo_atividade:
        m = re.search(r'\bTipo\s*(?:de\s*atividade)?\s*[,:]?\s*([^\n;]+?)(?=\s*Grupo\s|\s*Atividade\s*[,:]|\s*Data\s|Observa[cç][aã]o|\s*$)', texto_flat, re.IGNORECASE)
        if m:
            nc.tipo_atividade = _limpar_valor(m.group(1), 200)
    if not nc.grupo_atividade:
        m = re.search(r'\bGrupo\s*(?:de\s*atividade)?\s*[,:]?\s*([^\n;]+?)(?=\s*Tipo\s|\s*Atividade\s*[,:]|\s*Data\s|Observa[cç][aã]o|\s*$)', texto_flat, re.IGNORECASE)
        if m:
            nc.grupo_atividade = _limpar_valor(m.group(1), 100)

    # Por linha: "Tipo de Atividade: Faixa de Domínio" ou "Tipo Faixa de dominio" ou "Tipo, Segurança"
    for i, ln in enumerate(lines):
        m = re.match(r'Tipo\s*(?:de\s*)?[Aa]tividade\s*[,:]?\s*(.+)', ln, re.IGNORECASE)
        if m and m.group(1).strip():
            nc.tipo_atividade = _limpar_valor(m.group(1), 200)
            break
        m = re.match(r'Tipo\s*[,:]?\s*(.+)', ln, re.IGNORECASE)
        if m and m.group(1).strip() and not re.match(r'^(de\s+)?[Aa]tividade\s*[,:]?\s*$', m.group(1).strip(), re.IGNORECASE):
            nc.tipo_atividade = _limpar_valor(m.group(1), 200)
            break
        if re.match(r'Tipo\s*(?:de\s*)?[Aa]tividade\s*[,:]?\s*$', ln, re.IGNORECASE) and i + 1 < len(lines):
            nc.tipo_atividade = _limpar_valor(lines[i + 1], 200)
            break
        if re.match(r'Tipo\s*[,:]?\s*$', ln, re.IGNORECASE) and i + 1 < len(lines):
            nc.tipo_atividade = _limpar_valor(lines[i + 1], 200)
            break
    for i, ln in enumerate(lines):
        m = re.match(r'Grupo\s*(?:de\s*)?[Aa]tividade\s*[,:]?\s*(.+)', ln, re.IGNORECASE)
        if m and m.group(1).strip():
            nc.grupo_atividade = _limpar_valor(m.group(1), 100)
            break
        m = re.match(r'Grupo\s*[,:]?\s*(.+)', ln, re.IGNORECASE)
        if m and m.group(1).strip() and not re.match(r'^(de\s+)?[Aa]tividade\s*[,:]?\s*$', m.group(1).strip(), re.IGNORECASE):
            nc.grupo_atividade = _limpar_valor(m.group(1), 100)
            break
        if re.match(r'Grupo\s*(?:de\s*)?[Aa]tividade\s*[,:]?\s*$', ln, re.IGNORECASE) and i + 1 < len(lines):
            nc.grupo_atividade = _limpar_valor(lines[i + 1], 100)
            break
        if re.match(r'Grupo\s*[,:]?\s*$', ln, re.IGNORECASE) and i + 1 < len(lines):
            nc.grupo_atividade = _limpar_valor(lines[i + 1], 100)
            break

    # Fallback: "grupo" / "tipo" em qualquer posição da linha (ex.: "depressão atividade grupo pavimento tipo Pavimento")
    for ln in lines:
        if not nc.grupo_atividade:
            m = re.search(r'\bgrupo\b\s*(?:de\s*atividade)?\s*[,:]?\s*(.+?)(?=\s*tipo\b|\s*atividade\s*[,:]|\s*$)', ln, re.IGNORECASE)
            if m and m.group(1).strip():
                nc.grupo_atividade = _limpar_valor(m.group(1), 100)
        if not nc.tipo_atividade:
            m = re.search(r'\btipo\b\s*(?:de\s*atividade)?\s*[,:]?\s*(.+?)(?=\s*grupo\b|\s*atividade\s*[,:]|\s*$)', ln, re.IGNORECASE)
            if m and m.group(1).strip():
                nc.tipo_atividade = _limpar_valor(m.group(1), 200)
    for ln in lines:
        if not nc.grupo_atividade:
            m = re.search(r'\bgrupo\b\s*[,:]?\s*(.+)', ln, re.IGNORECASE)
            if m and m.group(1).strip():
                v = _limpar_valor(m.group(1), 200)
                if v and v.lower() not in ("de", "atividade"):
                    if " tipo " in (" " + v.lower() + " "):
                        v = v.split(" tipo ")[0].strip()
                    nc.grupo_atividade = v[:100]
                    break
        if not nc.tipo_atividade:
            m = re.search(r'\btipo\b\s*[,:]?\s*(.+)', ln, re.IGNORECASE)
            if m and m.group(1).strip():
                v = _limpar_valor(m.group(1), 200)
                if v and v.lower() not in ("de", "atividade"):
                    if " grupo " in (" " + v.lower() + " "):
                        v = v.split(" grupo ")[0].strip()
                    nc.tipo_atividade = v[:200]
                    break

    # ── observação (vários formatos: antes de "Observação:", depois de "Obs.:", etc.)
    for ln in lines:
        m = re.match(r'^(.+?)\s+Observa[cç][aã]o\s*:?\s*$', ln)
        if m:
            nc.observacao = m.group(1).strip()
            break
    if not nc.observacao:
        for ln in lines:
            m = re.search(r'Observa[cç][aã]o\s*:?\s*(.+)', ln, re.IGNORECASE)
            if m:
                nc.observacao = m.group(1).strip()
                break
    if not nc.observacao:
        for ln in lines:
            m = re.search(r'Obs\.?\s*:?\s*(.+)', ln, re.IGNORECASE)
            if m and len(m.group(1).strip()) > 0:
                nc.observacao = m.group(1).strip()
                break
    if not nc.observacao:
        for i, ln in enumerate(lines):
            if re.match(r'^Observa[cç][aã]o\s*:?\s*$', ln, re.IGNORECASE) or re.match(r'^Obs\.?\s*:?\s*$', ln, re.IGNORECASE):
                if i + 1 < len(lines) and lines[i + 1].strip():
                    nc.observacao = lines[i + 1].strip()[:200]
                    break
    partes_obs = []
    for ln in lines:
        m = re.search(r'Observa[cç][aã]o\s*:?\s*(.+)', ln, re.IGNORECASE)
        if m and m.group(1).strip():
            partes_obs.append(m.group(1).strip())
    if len(partes_obs) > 1:
        nc.observacao = " | ".join(partes_obs)[:500]
    elif len(partes_obs) == 1 and not nc.observacao:
        nc.observacao = partes_obs[0][:200]

    # ── rodovia SP (ex: "SP 075"); formato emergencial "RodoviaSP 075" (sem espaço)
    for ln in lines:
        m = re.search(r'Rodovia\s*\(SP\)\s*:?\s*(.+)', ln, re.IGNORECASE)
        if m and m.group(1).strip():
            nc.rodovia = m.group(1).strip()
            break
    if not nc.rodovia:
        m = re.search(r'Rodovia\s*SP\s*(\d+)', texto_flat, re.IGNORECASE)
        if m:
            nc.rodovia = "SP " + m.group(1).strip()
    if not nc.rodovia:
        m = re.search(r'RodoviaSP\s*(\d+)', texto_flat, re.IGNORECASE)
        if m:
            nc.rodovia = "SP " + m.group(1).strip()
    if not nc.rodovia:
        m = re.search(r'Rodovia\s*\(SP\)\s*:?\s*(\S+(?:\s+\S+)?)', texto_flat, re.IGNORECASE)
        if m:
            nc.rodovia = m.group(1).strip()
    if not nc.rodovia:
        for i, ln in enumerate(lines):
            if re.match(r'^Rodovia\s*\(SP\)\s*:?\s*$', ln, re.IGNORECASE):
                if i + 1 < len(lines) and lines[i + 1].strip():
                    nc.rodovia = lines[i + 1].strip().split()[0] if lines[i + 1].strip() else ""
                    if len(lines[i + 1].strip().split()) > 1:
                        nc.rodovia = lines[i + 1].strip()[:30]
                    break
    if not nc.rodovia:
        for i, ln in enumerate(lines):
            if re.match(r'^Rodovia\s*:?\s*$', ln, re.IGNORECASE):
                if i + 1 < len(lines) and re.search(r'SP\s*\d+', lines[i + 1], re.IGNORECASE):
                    nc.rodovia = lines[i + 1].strip()[:30]
                    break

    # ── prazo: data isolada APÓS a rodovia SP ou "Prazo" / "Data Limite" / "Data Programada Término" (emergencial)
    m = re.search(r'Data\s+Programada\s+T[eé]rmino\s*:?\s*(\d{2}/\d{2}/\d{4})', texto_flat, re.IGNORECASE)
    if m:
        nc.prazo_str = m.group(1)
    if not nc.prazo_str:
        m = re.search(r'Data\s+de\s+Execu[cç][aã]o\s+T[eé]rmino\s*:?\s*(\d{2}/\d{2}/\d{4})', texto_flat, re.IGNORECASE)
        if m:
            nc.prazo_str = m.group(1)
    if not nc.prazo_str:
        m = re.search(r'Prazo\s*:?\s*(\d{2}/\d{2}/\d{4})', texto_flat, re.IGNORECASE)
        if m:
            nc.prazo_str = m.group(1)
    if not nc.prazo_str:
        m = re.search(r'Data\s+Limite\s+(?:para\s+Reparo\s*)?:?\s*(\d{2}/\d{2}/\d{4})', texto_flat, re.IGNORECASE)
        if m:
            nc.prazo_str = m.group(1)
    if not nc.prazo_str:
        for i, ln in enumerate(lines):
            if re.search(r'Data\s+Programada\s+T[eé]rmino\s*:?\s*$', ln, re.IGNORECASE) or re.search(r'Data\s+de\s+Execu[cç][aã]o\s+T[eé]rmino\s*:?\s*$', ln, re.IGNORECASE):
                if i + 1 < len(lines) and re.match(r'^\d{2}/\d{2}/\d{4}', lines[i + 1]):
                    nc.prazo_str = lines[i + 1].strip()[:10]
                    break
    if not nc.prazo_str:
        for i, ln in enumerate(lines):
            if re.match(r'^Prazo\s*:?\s*$', ln, re.IGNORECASE) or re.search(r'Data\s+Limite\s+para\s+Reparo\s*$', ln, re.IGNORECASE):
                if i + 1 < len(lines) and re.match(r'^\d{2}/\d{2}/\d{4}', lines[i + 1]):
                    nc.prazo_str = lines[i + 1].strip()[:10]
                    break
    if not nc.prazo_str:
        rodsp_idx = next(
            (i for i, ln in enumerate(lines) if "Rodovia (SP):" in ln or re.search(r'Rodovia\s*\(SP\)', ln)), -1
        )
        if rodsp_idx >= 0:
            for ln in lines[rodsp_idx + 1: rodsp_idx + 5]:
                if re.match(r'^\d{2}/\d{2}/\d{4}$', ln.strip()):
                    nc.prazo_str = ln.strip()
                    break

    # ── rodovia nome
    for ln in lines:
        m = re.match(r'Rodovia:\s*(.+)', ln, re.IGNORECASE)
        if m:
            nc.rodovia_nome = m.group(1).strip()
            break

    # ── calcular prazo em dias e classificar
    if nc.data_con and nc.prazo_str:
        nc.prazo_dias = _prazo_dias(nc.data_con, nc.prazo_str)
        if nc.prazo_dias is not None:
            nc.emergencial = nc.prazo_dias <= PRAZO_EMERG_MAX

    nc.tipo_panela = _is_panela(nc.atividade)

    # Se KM final não foi extraído, usar o mesmo que KM inicial (evento localizado; extensão vem na NC quando houver)
    if (nc.km_ini_str or nc.km_ini) and (not nc.km_fim_str or nc.km_fim == 0.0):
        nc.km_fim_str = nc.km_ini_str or ""
        nc.km_fim = nc.km_ini

    # Fallback final: preencher E, O, P a partir de padrões soltos no bloco
    if not nc.horario_fiscalizacao:
        m = re.search(r'\b(\d{1,2}:\d{2}(?::\d{2})?)\b', block)
        if m:
            nc.horario_fiscalizacao = m.group(1)
    if not nc.tipo_atividade or not nc.grupo_atividade:
        VALORES_TIPO = ("Faixa de Domínio", "Faixa de dominio", "Pavimento", "Segurança", "Conservação")
        VALORES_GRUPO = ("Limpeza", "Pavimento", "Segurança", "Vedos", "Cercas Alambrados e Telamento", "Conservação")
        for i, ln in enumerate(lines):
            ln_clean = ln.strip()
            if not nc.tipo_atividade:
                for v in VALORES_TIPO:
                    if v.lower() in ln_clean.lower() and len(ln_clean) < 80:
                        if i > 0 and re.search(r'\btipo\b', lines[i - 1], re.IGNORECASE):
                            nc.tipo_atividade = ln_clean[:200]
                            break
                        if re.match(r'^' + re.escape(v) + r'\s*$', ln_clean, re.IGNORECASE):
                            nc.tipo_atividade = v[:200]
                            break
                if not nc.tipo_atividade and re.search(r'\b(Faixa\s+de\s+[Dd]om[ií]nio|Pavimento|Seguran[cç]a)\b', ln_clean):
                    nc.tipo_atividade = re.search(r'(Faixa\s+de\s+[Dd]om[ií]nio|Pavimento|Seguran[cç]a)', ln_clean, re.IGNORECASE).group(1).strip()[:200]
            if not nc.grupo_atividade:
                for v in VALORES_GRUPO:
                    if v.lower() in ln_clean.lower() and len(ln_clean) < 120:
                        if i > 0 and re.search(r'\bgrupo\b', lines[i - 1], re.IGNORECASE):
                            nc.grupo_atividade = ln_clean[:100]
                            break
                        if re.match(r'^' + re.escape(v) + r'\s*$', ln_clean, re.IGNORECASE):
                            nc.grupo_atividade = v[:100]
                            break
                if not nc.grupo_atividade and re.search(r'\b(Limpeza|Pavimento|Vedos|Cercas)', ln_clean, re.IGNORECASE):
                    m = re.search(r'((?:Vedos[,\s]*)?Cercas\s+Alambrados\s+e\s+Telamento|Limpeza|Pavimento|Seguran[cç]a)', ln_clean, re.IGNORECASE)
                    if m:
                        nc.grupo_atividade = m.group(1).strip()[:100]

    # Regra de negócio: pavimento → Pavimento (tipo e grupo); cerca/cercas → Segurança; Limpeza permanece no grupo
    texto_nc = " ".join(filter(None, [nc.atividade, nc.tipo_atividade, nc.grupo_atividade, block])).lower()
    if re.search(r'\b(pavimento|depress[aã]o|buraco|recapeamento|panela|reparo\s+de\s+pavimento)\b', texto_nc):
        nc.tipo_atividade = "Pavimento"
        nc.grupo_atividade = "Pavimento"
    elif re.search(r'\b(cerca|cercas|alambrado|telamento|vedos|reparo\s+de\s+cerca)\b', texto_nc):
        nc.tipo_atividade = "Segurança"
        nc.grupo_atividade = "Segurança"
    # Limpeza: grupo = Limpeza, tipo = Faixa de Domínio (não sobrescrever se já for Limpeza)
    if re.search(r'\b(limpeza|remo[cç][aã]o\s+de\s+lixo|faixa\s+de\s+dom[ií]nio)\b', texto_nc) and nc.grupo_atividade:
        if "limpeza" in (nc.grupo_atividade or "").lower():
            if not nc.tipo_atividade:
                nc.tipo_atividade = "Faixa de Domínio"

    # Grupo = resumo da atividade; Tipo = descrição do tipo. Inferir da atividade quando vazios.
    if (not nc.grupo_atividade or not nc.tipo_atividade) and nc.atividade:
        grupo_inf, tipo_inf = _inferir_grupo_tipo_da_atividade(nc.atividade)
        if grupo_inf and not nc.grupo_atividade:
            nc.grupo_atividade = grupo_inf
        if tipo_inf and not nc.tipo_atividade:
            nc.tipo_atividade = tipo_inf

    # NC válida precisa ao menos de código ou atividade
    if not (nc.codigo or nc.atividade):
        return None
    if logger.isEnabledFor(logging.DEBUG) and (not nc.tipo_atividade or not nc.grupo_atividade or not nc.horario_fiscalizacao):
        logger.debug(
            "NC block (tipo=%r grupo=%r horario=%r) text sample:\n%.800s",
            nc.tipo_atividade, nc.grupo_atividade, nc.horario_fiscalizacao, block
        )
    return nc


def _atribuir_grupo(nc: NcItem, mapa_eaf: list[dict]) -> None:
    """
    Atribui o número do Grupo EAF (coluna V do template) e o nome da empresa
    à NC com base no mapeamento por trecho (rodovia + km), conforme Contatos EAFs.
    Usa nc_artesp.utils.helpers.obter_grupo_empresa_por_trecho.
    """
    from nc_artesp.utils.helpers import obter_grupo_empresa_por_trecho
    grupo, empresa = obter_grupo_empresa_por_trecho(nc.rodovia, nc.km_ini, mapa_eaf)
    nc.grupo = grupo
    nc.empresa = empresa


def parse_pdf_nc(pdf_bytes: bytes) -> list[NcItem]:
    """Extrai NCs do PDF (Constatação/Código). Atributos NcItem = CAMPOS_TEMPLATE_LIST. Ordenado por rodovia, sentido, km_ini."""
    texto = _extrair_texto_pdf(pdf_bytes)

    # Remove cabeçalhos de página e rodapés
    texto = re.sub(r'Relat[oó]rio de Conserva[cç][aã]o de Rotina\s*', '', texto)
    texto = re.sub(r'--\s*\d+\s*of\s*\d+\s*--', '', texto)
    texto = re.sub(r'^\s*\d+\s*\d{2}/\d{2}/\d{4}\s*$', '', texto, flags=re.MULTILINE)

    # Divide nos blocos de NC (começa em data + "Constatação -")
    partes = re.split(r'(?=\d{2}/\d{2}/\d{4}\s+Constatação)', texto)

    ncs: list[NcItem] = []
    for bloco in partes:
        if "Constatação" not in bloco:
            continue
        nc = _parse_nc_block(bloco)
        if nc:
            ncs.append(nc)

    # PDF com legenda só "Código" (sem "Constatação" ao lado da data): divide por "Código" + número
    if len(ncs) == 0 and re.search(r'C[oó]digo\s*:?\s*\d{4,}', texto, re.IGNORECASE):
        partes_alt = re.split(r'(?=C[oó]digo\s*:?\s*\d{4,})', texto, flags=re.IGNORECASE)
        for bloco in partes_alt:
            bloco = bloco.strip()
            if not bloco or not re.search(r'\d{4,}', bloco):
                continue
            nc = _parse_nc_block(bloco)
            if nc:
                ncs.append(nc)
    # Layout emergencial: cada NC começa com linha só "Código" (valor em linha abaixo ou mais adiante)
    if len(ncs) == 0 and re.search(r'(?m)^C[oó]digo\s*$', texto, re.IGNORECASE):
        partes_alt = re.split(r'(?m)(?=^C[oó]digo\s*$)', texto, flags=re.IGNORECASE)
        for bloco in partes_alt:
            bloco = bloco.strip()
            if not bloco or not re.search(r'C[oó]digo', bloco, re.IGNORECASE):
                continue
            nc = _parse_nc_block(bloco)
            if nc:
                ncs.append(nc)

    # Atribui EAF/Grupo e ordena: Grupo → Rodovia → Sentido → KM
    for nc in ncs:
        _atribuir_grupo(nc, _MAPA_EAF_PADRAO)
    ncs.sort(key=lambda n: (n.grupo or 999, n.rodovia, n.sentido, n.km_ini))
    return ncs


# ─────────────────────────────────────────────────────────────────────────────
# ANÁLISE DE SEQUÊNCIA DE KM
# ─────────────────────────────────────────────────────────────────────────────

def _trecho_do_grupo_para_nc(nc: NcItem, mapa_eaf: list[dict]) -> tuple[float, float] | None:
    """
    Retorna (km_ini, km_fim) do trecho do MAPA_EAF que contém a NC (rodovia + km_ini).
    Só considera o grupo da própria NC. Retorna None se grupo não tiver trecho ou NC fora do trecho.
    """
    from nc_artesp.utils.helpers import normalizar_rodovia_para_busca
    g = nc.grupo or 0
    if not g or not mapa_eaf:
        return None
    rod_nc = normalizar_rodovia_para_busca(nc.rodovia)
    km = nc.km_ini
    for entry in mapa_eaf:
        if entry.get("grupo") != g:
            continue
        for trecho in entry.get("trechos", []):
            rod_t = normalizar_rodovia_para_busca(trecho.get("rodovia", ""))
            if not rod_t:
                continue
            if rod_t == rod_nc or rod_t in rod_nc or rod_nc in rod_t:
                ki = trecho.get("km_ini", 0.0)
                kf = trecho.get("km_fim", 9999.0)
                if ki <= km <= kf:
                    return (ki, kf)
    return None


def analisar_gaps(ncs: list[NcItem], limiar_km: float = LIMIAR_GAP_KM,
                  mapa_eaf: list[dict] | None = None) -> list[GapAlerta]:
    """
    Detecta saltos de KM acima do limiar apenas entre NCs do mesmo grupo, mesma data
    de constatação e mesmo trecho (rodovia + intervalo km) daquele grupo. Não mistura
    grupos, datas nem trechos. NCs de Meio Ambiente são excluídas.
    Retorna lista de GapAlerta ordenada por grupo → rodovia → sentido → km.
    """
    if not ncs:
        return []
    ncs = [nc for nc in ncs if not getattr(nc, "origem_ma", False)]
    mapa = mapa_eaf if mapa_eaf is not None else _MAPA_EAF_PADRAO

    # Só entram na análise NCs cujo (rodovia, km) está dentro de um trecho do seu grupo
    com_trecho: list[tuple[NcItem, tuple[float, float]]] = []
    for nc in ncs:
        t = _trecho_do_grupo_para_nc(nc, mapa)
        if t is not None:
            com_trecho.append((nc, t))

    # Agrupar por (grupo, data_con, rodovia, sentido, trecho_ini, trecho_fim)
    buckets: dict[tuple, list[NcItem]] = {}
    for nc, (trecho_ini, trecho_fim) in com_trecho:
        data_con = (nc.data_con or "").strip()
        chave = (nc.grupo or 0, data_con, nc.rodovia, nc.sentido, trecho_ini, trecho_fim)
        buckets.setdefault(chave, []).append(nc)

    alertas: list[GapAlerta] = []
    for (grupo_num, data_con, rodovia, sentido, _ti, _tf), bucket in sorted(buckets.items()):
        if len(bucket) < 2:
            continue
        bucket_ord = sorted(bucket, key=lambda n: n.km_ini)
        empresa = bucket_ord[0].empresa if bucket_ord else ""
        for i in range(len(bucket_ord) - 1):
            a = bucket_ord[i]
            b = bucket_ord[i + 1]
            gap = b.km_ini - a.km_fim
            if gap >= limiar_km:
                alertas.append(GapAlerta(
                    grupo=grupo_num,
                    empresa=empresa,
                    rodovia=rodovia,
                    sentido=sentido,
                    km_antes=a.km_fim,
                    km_depois=b.km_ini,
                    gap_km=round(gap, 3),
                    nc_antes=a.codigo,
                    nc_depois=b.codigo,
                ))

    return alertas


def analisar_sequencia_codigos(ncs: list[NcItem]) -> list[CodigoGapAlerta]:
    """
    Detecta saltos na numeração sequencial do Código Fiscalização por grupo EAF.
    Apenas NCs de Conservação (pavimento) entram na análise — Meio Ambiente
    não fiscaliza pavimento, portanto não gera alerta de apontamentos não entregues.

    O Código Fiscalização é atribuído sequencialmente pelo sistema ARTESP.
    Um salto (ex: 896643 → 896648) indica que os códigos 896644–896647 foram
    gerados mas NÃO foram entregues à concessionária.
    Esses apontamentos ocultos são potencialmente BURACOS/PANELAS NA PISTA
    — único tipo com prazo de 24 horas e que pode ser retido.
    """
    if not ncs:
        return []
    ncs = [nc for nc in ncs if not getattr(nc, "origem_ma", False)]

    def _para_int(codigo: str) -> Optional[int]:
        digits = re.sub(r'\D', '', str(codigo))
        return int(digits) if digits else None

    # Agrupar por grupo EAF
    buckets: dict[int, list[NcItem]] = {}
    for nc in ncs:
        g = nc.grupo or 0
        buckets.setdefault(g, []).append(nc)

    alertas: list[CodigoGapAlerta] = []
    for grupo_num, bucket in sorted(buckets.items()):
        # Filtrar NCs com código numérico válido e ordenar
        com_num = [(n, _para_int(n.codigo)) for n in bucket]
        com_num = [(nc, num) for nc, num in com_num if num is not None]
        if len(com_num) < 2:
            continue
        com_num.sort(key=lambda x: x[1])
        empresa = com_num[0][0].empresa or ""

        for i in range(len(com_num) - 1):
            nc_a, num_a = com_num[i]
            nc_b, num_b = com_num[i + 1]
            diff = num_b - num_a
            if diff <= 1:
                continue
            # Há códigos faltantes entre num_a e num_b
            faltantes = [str(num_a + j) for j in range(1, min(diff, 11))]
            alertas.append(CodigoGapAlerta(
                grupo=grupo_num,
                empresa=empresa,
                codigo_antes=nc_a.codigo,
                codigo_depois=nc_b.codigo,
                n_faltantes=diff - 1,
                codigos_faltantes=faltantes,
            ))

    return alertas


def resumo_estatistico(ncs: list[NcItem]) -> dict:
    """Retorna dicionário com métricas resumidas para o cabeçalho do relatório."""
    if not ncs:
        return {}
    tipos = {}
    for nc in ncs:
        tipos[nc.atividade] = tipos.get(nc.atividade, 0) + 1
    emergenciais = [nc for nc in ncs if nc.emergencial]
    panelas_poss = [nc for nc in ncs if nc.tipo_panela]
    rodovias     = sorted(set(nc.rodovia for nc in ncs if nc.rodovia))
    data_con     = ncs[0].data_con if ncs else ""
    # Resumo por grupo EAF (Conservação) e Meio Ambiente separado
    grupos: dict[int, dict] = {}
    for nc in ncs:
        if getattr(nc, "origem_ma", False):
            g = -1  # Meio Ambiente (exibido separado no resumo)
            emp = "Meio Ambiente"
        else:
            g = nc.grupo or 0
            emp = nc.empresa
        if g not in grupos:
            grupos[g] = {"grupo": g, "empresa": emp, "total": 0, "emergenciais": 0}
        grupos[g]["total"] += 1
        if nc.emergencial:
            grupos[g]["emergenciais"] += 1
    return {
        "total":          len(ncs),
        "tipos":          tipos,
        "n_tipos":        len(tipos),
        "emergenciais":   emergenciais,
        "panelas":        panelas_poss,
        "rodovias":       rodovias,
        "data_con":       data_con,
        "lote":           ncs[0].lote if ncs else "",
        "grupos":         dict(sorted(grupos.items())),
    }


# ─────────────────────────────────────────────────────────────────────────────
# GERAÇÃO DO RELATÓRIO PDF (ReportLab)
# ─────────────────────────────────────────────────────────────────────────────

# Paleta ARTESP
COR_HEADER    = colors.HexColor("#1e3a5f")   # azul escuro
COR_ALERTA    = colors.HexColor("#c0392b")   # vermelho NC
COR_EMERG     = colors.HexColor("#e74c3c")   # vermelho emergencial
COR_OK        = colors.HexColor("#27ae60")   # verde ok
COR_AVISO     = colors.HexColor("#e67e22")   # laranja aviso
COR_LINHAR    = colors.HexColor("#2c3e50")   # cabeçalho tabela
COR_LINHA_ALT = colors.HexColor("#f0f4f8")   # fundo linha par
COR_EMERG_BG  = colors.HexColor("#fdecea")   # fundo linha emergencial


def _estilos():
    ss = getSampleStyleSheet()
    extra = {
        "titulo": ParagraphStyle("titulo",
            fontName="Helvetica-Bold", fontSize=16,
            textColor=COR_HEADER, spaceAfter=4, alignment=TA_CENTER),
        "subtitulo": ParagraphStyle("subtitulo",
            fontName="Helvetica-Bold", fontSize=12,
            textColor=COR_HEADER, spaceAfter=2, alignment=TA_CENTER),
        "secao": ParagraphStyle("secao",
            fontName="Helvetica-Bold", fontSize=11,
            textColor=colors.white, spaceAfter=0, leading=16),
        "corpo": ParagraphStyle("corpo",
            fontName="Helvetica", fontSize=9,
            textColor=colors.HexColor("#2c3e50"), spaceAfter=2, leading=13),
        "alerta": ParagraphStyle("alerta",
            fontName="Helvetica-Bold", fontSize=9,
            textColor=COR_ALERTA, spaceAfter=2, leading=13),
        "emerg": ParagraphStyle("emerg",
            fontName="Helvetica-Bold", fontSize=9,
            textColor=COR_EMERG, spaceAfter=2, leading=13),
        "tabcab": ParagraphStyle("tabcab",
            fontName="Helvetica-Bold", fontSize=8,
            textColor=colors.white, alignment=TA_CENTER),
        "tabcel": ParagraphStyle("tabcel",
            fontName="Helvetica", fontSize=8,
            textColor=colors.HexColor("#2c3e50"), alignment=TA_LEFT),
        "tabcel_emerg": ParagraphStyle("tabcel_emerg",
            fontName="Helvetica-Bold", fontSize=8,
            textColor=COR_EMERG, alignment=TA_LEFT),
        "rodape": ParagraphStyle("rodape",
            fontName="Helvetica", fontSize=7,
            textColor=colors.grey, alignment=TA_CENTER),
        "celula": ParagraphStyle("celula",
            fontName="Helvetica", fontSize=9,
            textColor=colors.HexColor("#2c3e50"), spaceAfter=0, leading=11,
            alignment=TA_LEFT),
    }
    return extra


def _km_fmt(v: float) -> str:
    km  = int(v)
    met = round((v - km) * 1000)
    return f"{km}+{met:03d}"


def _safe_latin1(s: str) -> str:
    """Garante string encodável em Latin-1 (evita erro em ReportLab/Helvetica e headers)."""
    if not s:
        return s
    nfd = unicodedata.normalize("NFD", s)
    sem_comb = "".join(c for c in nfd if unicodedata.category(c) != "Mn")
    try:
        return sem_comb.encode("latin-1").decode("latin-1")
    except UnicodeEncodeError:
        return sem_comb.encode("latin-1", "replace").decode("latin-1")


def _banner(texto: str, cor: colors.Color, estilos: dict):
    """Faixa colorida com texto branco — usada como cabeçalho de seção."""
    return Table(
        [[Paragraph(_safe_latin1(texto), estilos["secao"])]],
        colWidths=[170 * mm],
        style=TableStyle([
            ("BACKGROUND", (0, 0), (-1, -1), cor),
            ("LEFTPADDING",  (0, 0), (-1, -1), 8),
            ("RIGHTPADDING", (0, 0), (-1, -1), 8),
            ("TOPPADDING",   (0, 0), (-1, -1), 6),
            ("BOTTOMPADDING",(0, 0), (-1, -1), 6),
        ]),
    )


def _tabela_ncs(ncs_tipo: list[NcItem], estilos: dict) -> Table:
    """Tabela com os dados originais de cada NC (sem alteração)."""
    cabecalho = ["Cód. Fiscal.", "Grp", "KM Inicial", "KM Final", "Sentido",
                 "Data Const.", "Prazo", "Obs"]
    linhas = [
        [Paragraph(c, estilos["tabcab"]) for c in cabecalho]
    ]
    colunas_w = [26*mm, 10*mm, 20*mm, 20*mm, 26*mm, 20*mm, 20*mm, 28*mm]

    for nc in sorted(ncs_tipo, key=lambda n: n.km_ini):
        est = estilos["tabcel_emerg"] if nc.emergencial else estilos["tabcel"]
        prazo_txt = nc.prazo_str + (" ⚠" if nc.emergencial else "")
        grupo_txt = str(nc.grupo) if nc.grupo else "—"
        linhas.append([
            Paragraph(nc.codigo,      est),
            Paragraph(grupo_txt,      est),
            Paragraph(_km_fmt(nc.km_ini) if nc.km_ini else nc.km_ini_str, est),
            Paragraph(_km_fmt(nc.km_fim) if nc.km_fim else nc.km_fim_str, est),
            Paragraph(nc.sentido,     est),
            Paragraph(nc.data_con,    est),
            Paragraph(prazo_txt,      est),
            Paragraph(nc.observacao[:40] if nc.observacao else "—", est),
        ])

    ts = TableStyle([
        # Cabeçalho
        ("BACKGROUND",   (0, 0), (-1, 0), COR_LINHAR),
        ("GRID",         (0, 0), (-1, -1), 0.3, colors.HexColor("#bdc3c7")),
        ("TOPPADDING",   (0, 0), (-1, -1), 3),
        ("BOTTOMPADDING",(0, 0), (-1, -1), 3),
        ("LEFTPADDING",  (0, 0), (-1, -1), 5),
        ("RIGHTPADDING", (0, 0), (-1, -1), 5),
        ("VALIGN",       (0, 0), (-1, -1), "MIDDLE"),
    ])
    # Fundo alternado e destaque emergencial
    for i, nc in enumerate(ncs_tipo, start=1):
        row = i
        if nc.emergencial:
            ts.add("BACKGROUND", (0, row), (-1, row), COR_EMERG_BG)
        elif i % 2 == 0:
            ts.add("BACKGROUND", (0, row), (-1, row), COR_LINHA_ALT)

    return Table(linhas, colWidths=colunas_w, style=ts, repeatRows=1)


def _km_para_partes(km: float) -> tuple[int, int]:
    """Converte km decimal em (km_int, metros). Ex.: 103.5 → (103, 500)."""
    if not km and km != 0.0:
        return 0, 0
    k = int(km)
    m = round((km - k) * 1000)
    if m >= 1000:
        m = 0
    return k, m


def _lote_num_do_pdf(nc: "NcItem") -> str | None:
    """Extrai o número do lote do item (ex.: '13', '21', '50'). Só retorna se for lote conhecido (evita confundir com código da fiscalização)."""
    try:
        from nc_artesp.config import _LOTE_CONCESSIONARIA
        lotes_conhecidos = set(_LOTE_CONCESSIONARIA.keys()) if _LOTE_CONCESSIONARIA else set()
    except Exception:
        lotes_conhecidos = set()
    for s in [(nc.lote or "").strip(), (nc.concessionaria or "").strip()]:
        if not s:
            continue
        if re.match(r"^\d+$", s) and s in lotes_conhecidos:
            return s
        for num in re.findall(r"\d+", s):
            if num in lotes_conhecidos:
                return num
    return None


def _concessionaria_por_lote(lote_ou_concessionaria: str) -> str:
    """Retorna 'Lote N Nome da Concessionária' a partir do número do lote (ex.: '13' → 'Lote 13 Rodovias das Colinas')."""
    try:
        from nc_artesp.config import _LOTE_CONCESSIONARIA
        s = (lote_ou_concessionaria or "").strip()
        n = re.search(r"\d+", s)
        if n:
            num = n.group(0)
            nome = _LOTE_CONCESSIONARIA.get(num)
            if nome:
                return f"Lote {num} {nome}"
    except Exception:
        pass
    return (lote_ou_concessionaria or "").strip()


def _caminho_template_relatorio_xlsx() -> Path:
    """Caminho do template do relatório (config: ARTESP_TEMPLATE_RELATORIO ou assets/templates)."""
    from nc_artesp.config import TEMPLATE_RELATORIO_XLSX
    return Path(TEMPLATE_RELATORIO_XLSX).resolve()


# Colunas do template de saída (relatório XLSX). km_ini: col 8+9; km_fim: 10+11.
COLUNAS_TEMPLATE_SAIDA: dict[str, int] = {
    "codigo": 3, "data_con": 4, "horario_fiscalizacao": 5, "rodovia": 6, "concessionaria": 7,
    "km_ini_str": 8, "km_fim_str": 10, "sentido": 12, "tipo_atividade": 15, "grupo_atividade": 16,
    "atividade": 17, "prazo_str": 19, "empresa": 20, "nome_fiscal": 21,
}


def _detectar_colunas_saida_template(ws, cabecalho_fim: int = 4) -> dict[str, int]:
    """Detecta coluna Responsável Técnico no cabeçalho do template de saída."""
    col_map = dict(COLUNAS_TEMPLATE_SAIDA)
    for r in range(1, min(cabecalho_fim + 1, ws.max_row + 1)):
        for c in range(1, min(ws.max_column + 1, 30)):
            v = ws.cell(row=r, column=c).value
            s = (str(v or "")).strip().lower().replace("é", "e").replace("á", "a")
            if v and ("responsavel" in s or "responsável" in s or ("respons" in s and "tecnico" in s)):
                col_map["nome_fiscal"] = c
                return col_map
    return col_map


def gerar_relatorio_xlsx(ncs: list[NcItem], lote_selecionado: str | None = None) -> bytes:
    """Preenche o template XLSX a partir de NcItem. Coluna fica vazia só quando a informação não está nos documentos lidos."""
    if not OPENPYXL_OK:
        raise ImportError("openpyxl não instalado: pip install openpyxl")
    template_path = _caminho_template_relatorio_xlsx()
    CABECALHO_FIM = 4
    PRIMEIRA_LINHA_DADOS = 5
    wb = None
    ws = None

    if template_path.is_file():
        logger.info("Relatório XLSX: usando template %s", template_path)
        try:
            if template_path.suffix.lower() == ".xls":
                import xlrd
                raw = template_path.read_bytes()
                rb = xlrd.open_workbook(file_contents=raw)
                sh = rb.sheet_by_index(0)
                wb = openpyxl.Workbook()
                ws = wb.active
                ws.title = "Relatório"
                for r in range(min(CABECALHO_FIM, sh.nrows)):
                    for c in range(min(sh.ncols, 30)):
                        v = sh.cell_value(r, c)
                        ws.cell(row=r + 1, column=c + 1, value=v)
            else:
                import shutil
                import tempfile
                fd, tmp_path = tempfile.mkstemp(suffix=".xlsx")
                try:
                    os.close(fd)
                    shutil.copy2(str(template_path), tmp_path)
                    wb = openpyxl.load_workbook(tmp_path)
                    ws = wb.active
                    while ws.max_row >= PRIMEIRA_LINHA_DADOS:
                        ws.delete_rows(ws.max_row, 1)
                finally:
                    try:
                        Path(tmp_path).unlink(missing_ok=True)
                    except Exception:
                        pass
        except Exception as e:
            logger.exception("Template relatório %s: %s", template_path, e)
            raise FileNotFoundError(
                f"Template do relatório não carregou: {template_path}\nErro: {e}"
            ) from e
    else:
        logger.warning(
            "Template relatório não encontrado (path=%s); modo fallback com cabeçalho padrão.",
            template_path,
        )

    if wb is None or ws is None:
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Relatório"
        for col, val in enumerate([
            "", "", "Cód. Fiscalização", "Data Fiscalização", "Horário da Fiscalização",
            "Rodovia", "Concessionária Lote", "Trecho", "", "", "", "Sentido",
            "Data Retorno", "Status Retorno", "Tipo de Atividade", "Grupo de Atividade",
            "Atividade", "Data Envio", "Data Reparo", "EAF", "Responsável Técnico",
        ], start=1):
            if col <= 21 and val:
                ws.cell(row=2, column=col, value=val)
        for col, val in enumerate(["", ""] + [""] * 5 + [" Km Inicial", "m", "Km Final", "m", ""] + [""] * 8 + ["Responsável Técnico"], start=1):
            if col <= 21 and val:
                ws.cell(row=3, column=col, value=val)

    col_map = COLUNAS_TEMPLATE_SAIDA.copy()
    if wb is not None and ws is not None and hasattr(ws, "max_column"):
        try:
            col_map = _detectar_colunas_saida_template(ws, CABECALHO_FIM)
        except Exception:
            pass

    for row_idx, nc in enumerate(ncs, start=PRIMEIRA_LINHA_DADOS):
        km_ini_k, km_ini_m = _km_para_partes(nc.km_ini)
        km_fim_k, km_fim_m = _km_para_partes(nc.km_fim)
        conc_val = _concessionaria_por_lote(lote_selecionado) if lote_selecionado else _concessionaria_por_lote(nc.concessionaria or nc.lote) or (nc.concessionaria or nc.lote or "").strip()
        # Preencher quando a informação existe nos documentos; vazio só quando não presente
        ws.cell(row=row_idx, column=col_map["codigo"], value=(nc.codigo or "").strip())
        ws.cell(row=row_idx, column=col_map["data_con"], value=(nc.data_con or "").strip())
        ws.cell(row=row_idx, column=col_map["horario_fiscalizacao"], value=(nc.horario_fiscalizacao or "").strip())
        ws.cell(row=row_idx, column=col_map["rodovia"], value=(nc.rodovia or "").strip())
        ws.cell(row=row_idx, column=col_map["concessionaria"], value=conc_val)
        ws.cell(row=row_idx, column=col_map["km_ini_str"], value=km_ini_k)
        ws.cell(row=row_idx, column=col_map["km_ini_str"] + 1, value=km_ini_m)
        ws.cell(row=row_idx, column=col_map["km_fim_str"], value=km_fim_k)
        ws.cell(row=row_idx, column=col_map["km_fim_str"] + 1, value=km_fim_m)
        ws.cell(row=row_idx, column=col_map["sentido"], value=(nc.sentido or "").strip())
        ws.cell(row=row_idx, column=13, value="")
        ws.cell(row=row_idx, column=14, value="")
        ws.cell(row=row_idx, column=col_map["tipo_atividade"], value=(nc.tipo_atividade or "").strip())
        ws.cell(row=row_idx, column=col_map["grupo_atividade"], value=(nc.grupo_atividade or "").strip())
        ws.cell(row=row_idx, column=col_map["atividade"], value=(nc.atividade or "").strip())
        ws.cell(row=row_idx, column=18, value=(nc.data_con or "").strip())
        ws.cell(row=row_idx, column=col_map["prazo_str"], value=(nc.prazo_str or "").strip())
        ws.cell(row=row_idx, column=col_map["empresa"], value=(nc.empresa or "").strip())
        ws.cell(row=row_idx, column=col_map["nome_fiscal"], value=(nc.nome_fiscal or "").strip())
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


def gerar_relatorio_pdf(ncs: list[NcItem],
                        alertas_km: list[GapAlerta],
                        alertas_codigo: list[CodigoGapAlerta],
                        limiar_km: float = LIMIAR_GAP_KM,
                        mapa_eaf: list | None = None) -> bytes:
    """Resumo geral em PDF. mapa_eaf = mapa do lote (trechos por grupo); se None, usa config padrão."""
    if not REPORTLAB_OK:
        raise ImportError("reportlab não instalado: pip install reportlab")
    mapa_uso = (mapa_eaf or _MAPA_EAF_PADRAO)

    def _rodape_pagina(canvas, doc):
        canvas.saveState()
        canvas.setFont("Helvetica", 8)
        w, _ = doc.pagesize
        canvas.drawCentredString(w / 2, 8 * mm, "Desenvolvedor Ozeias Engler")
        canvas.restoreState()

    est  = _estilos()
    buf  = io.BytesIO()
    doc  = SimpleDocTemplate(
        buf, pagesize=A4,
        leftMargin=18*mm, rightMargin=18*mm,
        topMargin=12*mm,  bottomMargin=12*mm,
        onFirstPage=_rodape_pagina,
        onLaterPages=_rodape_pagina,
    )
    story = []
    agora = datetime.now().strftime("%d/%m/%Y %H:%M")
    try:
        data_relatorio = datetime.strptime(agora, "%d/%m/%Y %H:%M").date()
    except ValueError:
        data_relatorio = date.today()

    res    = resumo_estatistico(ncs)
    emerg  = res.get("emergenciais", [])
    rodovs = ", ".join(res.get("rodovias", [])) or "—"
    data_c = res.get("data_con", "")
    # Alerta de emergenciais (prazo 24 h) só vale para relatório do mesmo dia da constatação
    data_con_dt = _parse_data(data_c)
    data_con_date = data_con_dt.date() if data_con_dt else None
    relatorio_hoje = data_con_date is not None and data_relatorio == data_con_date
    lote   = res.get("lote", "")

    # ── CAPA ────────────────────────────────────────────────────────────────
    story.append(Spacer(1, 4*mm))
    story.append(Paragraph("ARTESP — Relatório de Análise de NCs", est["titulo"]))
    story.append(Paragraph("Conservação de Rotina", est["subtitulo"]))
    story.append(Spacer(1, 3*mm))
    story.append(HRFlowable(width="100%", thickness=2, color=COR_HEADER))
    story.append(Spacer(1, 3*mm))

    # Metadados — Paragraph em cada célula para quebra de texto e evitar overflow
    def _cel(s: str, bold: bool = False) -> Paragraph:
        style = ParagraphStyle("cel", parent=est["celula"], fontName="Helvetica-Bold" if bold else "Helvetica")
        t = _safe_latin1(s or "").replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
        return Paragraph(t, style)
    meta = [
        [_cel("Rodovia(s):", True), _cel(rodovs),   _cel("Lote:", True), _cel(lote)],
        [_cel("Data Constatação:", True), _cel(data_c), _cel("Emitido em:", True), _cel(agora)],
        [_cel("Total de NCs:", True), _cel(str(res.get("total", 0))),
         _cel("Tipos únicos:", True), _cel(str(res.get("n_tipos", 0)))],
        [_cel("NCs emergenciais (24 h):", True), _cel(str(len(emerg))),
         _cel("Alertas de salto de km:", True), _cel(str(len(alertas_km)))],
        [_cel("Gaps de km:", True),
         _cel(str(sum(a.n_faltantes for a in alertas_codigo))),
         _cel("Grupos com ocultação:", True), _cel(str(len(alertas_codigo)))],
    ]
    meta_t = Table(meta, colWidths=[48*mm, 48*mm, 38*mm, 38*mm],
        style=TableStyle([
            ("TOPPADDING",    (0, 0), (-1, -1), 3),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
            ("LEFTPADDING",   (0, 0), (-1, -1), 5),
            ("RIGHTPADDING",  (0, 0), (-1, -1), 5),
            ("GRID", (0, 0), (-1, -1), 0.3, colors.HexColor("#bdc3c7")),
            ("BACKGROUND", (0, 0), (-1, -1), COR_LINHA_ALT),
        ]))
    story.append(meta_t)
    story.append(Spacer(1, 6*mm))

    # Resumo por Grupo EAF (equipes por trecho km — Contatos EAFs)
    def _trecho_resumo(grupo_num: int) -> str:
        """Texto rodovia e trecho coberto para o grupo (ex.: SP 75 km 15 > 50)."""
        if grupo_num is None or grupo_num == -1 or grupo_num == 0 or grupo_num == 999:
            return "—"
        for entry in mapa_uso:
            if entry.get("grupo") == grupo_num:
                partes = []
                for t in entry.get("trechos", []):
                    rod = t.get("rodovia", "").strip()
                    ki = t.get("km_ini", 0.0)
                    kf = t.get("km_fim", 0.0)
                    if rod:
                        km_ini_int = int(ki) if ki == int(ki) else ki
                        km_fim_int = int(kf) if kf == int(kf) else kf
                        partes.append(f"{rod} km {km_ini_int} > {km_fim_int}")
                return " | ".join(partes) if partes else "—"
        return "—"

    story.append(_banner("RESUMO POR GRUPO DE FISCALIZAÇÃO", COR_HEADER, est))
    story.append(Spacer(1, 2*mm))
    grupos_res = res.get("grupos", {})
    grp_dados = [[Paragraph("Grupo", est["tabcab"]), Paragraph("Rodovia / Trecho", est["tabcab"]),
                  Paragraph("Empresa", est["tabcab"]), Paragraph("NCs", est["tabcab"]), Paragraph("Emergenciais", est["tabcab"])]]
    if grupos_res:
        for g_num, g_info in sorted(grupos_res.items()):
            if g_num == -1:
                label = "Meio Ambiente"
            elif g_num:
                label = str(g_num)
            else:
                label = "Não ident."
            # Escapar só < e > para Paragraph; não escapar & para não quebrar ">" (evita &amp;gt; no PDF)
            trecho_txt = _trecho_resumo(g_num).replace("<", "&lt;").replace(">", "&gt;")
            emp = (g_info.get("empresa") or "—").replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
            grp_dados.append([
                Paragraph(label, est["celula"]),
                Paragraph(trecho_txt, est["celula"]),
                Paragraph(emp, est["celula"]),
                Paragraph(str(g_info.get("total", 0)), est["celula"]),
                Paragraph(str(g_info.get("emergenciais", 0)) or "0", est["celula"]),
            ])
    else:
        grp_dados.append([Paragraph("—", est["celula"]), Paragraph("—", est["celula"]),
                          Paragraph("Nenhum grupo identificado", est["celula"]),
                          Paragraph("0", est["celula"]), Paragraph("0", est["celula"])])
    grp_t = Table(grp_dados, colWidths=[28*mm, 52*mm, 52*mm, 16*mm, 26*mm],
        style=TableStyle([
            ("BACKGROUND",    (0, 0), (-1, 0), COR_LINHAR),
            ("TOPPADDING",    (0, 0), (-1, -1), 4),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
            ("LEFTPADDING",   (0, 0), (-1, -1), 6),
            ("RIGHTPADDING",  (0, 0), (-1, -1), 6),
            ("GRID",          (0, 0), (-1, -1), 0.3, colors.HexColor("#bdc3c7")),
            *[("BACKGROUND", (0, i), (-1, i), COR_LINHA_ALT)
              for i in range(2, len(grp_dados), 2)],
        ]))
    story.append(grp_t)
    story.append(Spacer(1, 4*mm))

    # Resumo de tipos
    story.append(_banner("RESUMO POR TIPO DE NC", COR_HEADER, est))
    story.append(Spacer(1, 2*mm))
    tipos_sorted = sorted(res.get("tipos", {}).items(), key=lambda x: -x[1])
    tipo_dados   = [[Paragraph("Atividade", est["tabcab"]), Paragraph("Qtd", est["tabcab"])]]
    for tipo, qtd in tipos_sorted:
        t_esc = (tipo or "").replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
        tipo_dados.append([Paragraph(t_esc, est["celula"]), Paragraph(str(qtd), est["celula"])])
    tipo_t = Table(tipo_dados, colWidths=[142*mm, 22*mm],
        style=TableStyle([
            ("BACKGROUND",    (0, 0), (-1, 0), COR_LINHAR),
            ("TOPPADDING",    (0, 0), (-1, -1), 4),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
            ("LEFTPADDING",   (0, 0), (-1, -1), 6),
            ("RIGHTPADDING",  (0, 0), (-1, -1), 6),
            ("GRID",          (0, 0), (-1, -1), 0.3, colors.HexColor("#bdc3c7")),
            *[("BACKGROUND", (0, i), (-1, i), COR_LINHA_ALT)
              for i in range(2, len(tipo_dados), 2)],
        ]))
    story.append(tipo_t)

    # ── Alertas: só para relatório do mesmo dia da constatação; anteriores entram só no resumo ──
    if relatorio_hoje:
        # ── ALERTAS DE APONTAMENTOS NÃO ENTREGUES (gap de código) ──────────────────
        if alertas_codigo:
            story.append(Spacer(1, 6*mm))
            story.append(_banner(
                f"🔴  APONTAMENTOS NÃO ENTREGUES — SALTO NA NUMERAÇÃO DO CÓDIGO",
                COR_EMERG, est
            ))
            story.append(Spacer(1, 2*mm))
            story.append(Paragraph(
                "O Código Fiscalização é atribuído <b>sequencialmente</b> pelo sistema ARTESP. "
                "Um salto na sequência indica apontamentos que foram gerados mas "
                "<b>NÃO foram entregues</b> à concessionária. "
                "Somente <b>buracos/panelas na pista</b> são retidos (prazo 24 h) — "
                "todos os demais tipos são sempre entregues.",
                est["corpo"]
            ))
            story.append(Spacer(1, 2*mm))

            total_ocultos = sum(a.n_faltantes for a in alertas_codigo)
            story.append(Paragraph(
                f"<font color='#e74c3c'><b>Total de apontamentos não entregues: "
                f"{total_ocultos}</b></font>",
                est["emerg"]
            ))
            story.append(Spacer(1, 3*mm))

            for i, ca in enumerate(alertas_codigo, 1):
                label_grp = f"Grupo {ca.grupo} — {ca.empresa}" if ca.grupo else "Grupo não identificado"
                faltantes_str = ", ".join(ca.codigos_faltantes)
                if ca.n_faltantes > 10:
                    faltantes_str += f" ... (+{ca.n_faltantes - 10} ocultos)"
                story.append(KeepTogether([
                    Paragraph(
                        f"<b>Ocorrência {i}</b> | {label_grp}",
                        est["emerg"]
                    ),
                    Paragraph(
                        f"Último entregue: <b>{ca.codigo_antes}</b> → "
                        f"Próximo entregue: <b>{ca.codigo_depois}</b> "
                        f"<font color='#e74c3c'><b>({ca.n_faltantes} não entregue(s))</b></font>",
                        est["corpo"]
                    ),
                    Paragraph(
                        f"Códigos ausentes: {faltantes_str}",
                        est["corpo"]
                    ),
                    Paragraph(
                        "⚠ Esses apontamentos são potencialmente BURACOS NA PISTA "
                        "(único tipo com prazo de 24 h que pode ser retido).",
                        est["emerg"]
                    ),
                    Spacer(1, 4*mm),
                ]))

        # ── ALERTAS DE SALTO DE KM ──────────────────────────────────────────────
        if alertas_km:
            story.append(Spacer(1, 4*mm))
            story.append(_banner(
                f"⚠  ALERTAS DE SALTO DE KM — TRECHO SEM APONTAMENTO (limiar: {limiar_km:.1f} km)",
                COR_AVISO, est
            ))
            story.append(Spacer(1, 2*mm))
            story.append(Paragraph(
                "Os trechos abaixo, dentro de cada equipe de fiscalização, "
                "apresentam intervalo superior ao limiar sem apontamentos. "
                "Verifique se há buracos/panelas não registrados nesse intervalo.",
                est["corpo"]
            ))
            story.append(Spacer(1, 2*mm))

            for i, ga in enumerate(alertas_km, 1):
                label_grp = f"Grupo {ga.grupo} — {ga.empresa}" if ga.grupo else "Grupo não identificado"
                story.append(KeepTogether([
                    Paragraph(
                        f"<b>Alerta {i}</b> | {label_grp} | {ga.rodovia} — Sentido: {ga.sentido}",
                        est["alerta"]
                    ),
                    Paragraph(
                        f"Trecho sem NC: km <b>{_km_fmt(ga.km_antes)}</b> → "
                        f"km <b>{_km_fmt(ga.km_depois)}</b> "
                        f"<font color='#c0392b'><b>(gap: {ga.gap_km:.3f} km)</b></font>",
                        est["corpo"]
                    ),
                    Paragraph(
                        f"NC antes: {ga.nc_antes} | NC depois: {ga.nc_depois}",
                        est["corpo"]
                    ),
                    Spacer(1, 3*mm),
                ]))

        # ── NCs EMERGENCIAIS ────────────────────────────────────────────────────
        if emerg:
            story.append(Spacer(1, 4*mm))
            story.append(_banner("🚨  NCs EMERGENCIAIS — PRAZO ≤ 24 h", COR_EMERG, est))
            story.append(Spacer(1, 2*mm))
            story.append(Paragraph(
                "As NCs abaixo têm prazo igual ou anterior a 1 dia após a constatação. "
                "Data e hora do prazo: até 23:59 do dia indicado.",
                est["corpo"]
            ))
            story.append(Spacer(1, 2*mm))
            for nc in sorted(emerg, key=lambda n: (n.rodovia, n.km_ini)):
                prazo_data_hora = f"{nc.prazo_str} 23:59" if nc.prazo_str else "—"
                story.append(KeepTogether([
                    Paragraph(
                        f"<b>Cód. {nc.codigo}</b> | {nc.rodovia} | "
                        f"km {_km_fmt(nc.km_ini)} → {_km_fmt(nc.km_fim)} | "
                        f"Sentido: {nc.sentido}",
                        est["emerg"]
                    ),
                    Paragraph(
                        f"Atividade: <b>{nc.atividade}</b>",
                        est["corpo"]
                    ),
                    Paragraph(
                        f"Prazo: <font color='#e74c3c'><b>{prazo_data_hora}</b></font> "
                        f"(24 h após constatação — {'MESMO DIA' if nc.prazo_dias == 0 else str(nc.prazo_dias) + ' dia(s)'})",
                        est["corpo"]
                    ),
                    *(
                        [Paragraph(f"Obs: {nc.observacao}", est["corpo"])]
                        if nc.observacao else []
                    ),
                    Spacer(1, 3*mm),
                ]))

    # ── NCs POR GRUPO EAF → TIPO → KM (dados originais intactos) ───────────────
    story.append(PageBreak())
    story.append(_banner(
        "NCs POR GRUPO DE FISCALIZAÇÃO — TIPO — KM  (dados originais)",
        COR_HEADER, est
    ))
    story.append(Spacer(1, 2*mm))
    story.append(Paragraph(
        "Dados reproduzidos fielmente do PDF original, sem qualquer alteração. "
        "Organizados por Grupo/EAF, depois por atividade "
        "e ordenados por KM crescente.",
        est["corpo"]
    ))
    story.append(Spacer(1, 4*mm))

    # Agrupar: Conservação por (grupo_num, empresa); Meio Ambiente em bloco à parte
    # Lista (label, grupo_num ou None para MA, por_tipo)
    blocos: list[tuple[str, int | None, dict[str, list[NcItem]]]] = []
    conservacao: dict[tuple, dict[str, list[NcItem]]] = {}
    meio_ambiente: dict[str, list[NcItem]] = {}
    for nc in ncs:
        if getattr(nc, "origem_ma", False):
            meio_ambiente.setdefault(nc.atividade or "—", []).append(nc)
        else:
            chave = (nc.grupo or 999, nc.empresa or "Sem EAF identificada")
            conservacao.setdefault(chave, {})
            conservacao[chave].setdefault(nc.atividade or "—", []).append(nc)
    for (grupo_num, empresa), por_tipo in sorted(conservacao.items()):
        label = f"GRUPO {grupo_num} — {empresa}" if grupo_num != 999 else "GRUPO NÃO IDENTIFICADO"
        blocos.append((label, grupo_num, por_tipo))
    if meio_ambiente:
        blocos.append(("MEIO AMBIENTE", None, meio_ambiente))

    for label_grupo, grupo_num, por_tipo in blocos:
        # Banner de grupo
        story.append(Table(
            [[Paragraph(label_grupo, ParagraphStyle(
                "ghdr", fontName="Helvetica-Bold", fontSize=13,
                textColor=colors.white, leading=18,
            ))]],
            colWidths=[170 * mm],
            style=TableStyle([
                ("BACKGROUND",    (0, 0), (-1, -1), COR_HEADER),
                ("TOPPADDING",    (0, 0), (-1, -1), 8),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 8),
                ("LEFTPADDING",   (0, 0), (-1, -1), 10),
            ])
        ))
        story.append(Spacer(1, 3*mm))

        # Trechos fiscalizados (só Conservação; Meio Ambiente não usa MAPA_EAF)
        if grupo_num is not None and grupo_num != 999:
            for entry in mapa_uso:
                if entry.get("grupo") == grupo_num:
                    partes = [
                        f"{t['rodovia']} km {t['km_ini']:.3f}–{t['km_fim']:.3f}"
                        for t in entry.get("trechos", [])
                    ]
                    story.append(Paragraph(
                        "<b>Trechos fiscalizados:</b> " + " | ".join(partes),
                        est["corpo"]
                    ))
                    break
        story.append(Spacer(1, 2*mm))

        # NCs por tipo dentro do grupo
        for idx_t, (tipo, grupo_ncs) in enumerate(sorted(por_tipo.items()), 1):
            tem_emerg = any(n.emergencial for n in grupo_ncs)
            cor_tipo  = COR_EMERG if tem_emerg else COR_ALERTA
            label_emg = "  🚨" if tem_emerg else ""
            story.append(KeepTogether([
                _banner(
                    f"{idx_t}. {tipo.upper()}  —  {len(grupo_ncs)} NC(s){label_emg}",
                    cor_tipo, est
                ),
                Spacer(1, 2*mm),
                _tabela_ncs(grupo_ncs, est),
                Spacer(1, 5*mm),
            ]))
        story.append(Spacer(1, 4*mm))

    # Rodapé final
    story.append(Spacer(1, 4*mm))
    story.append(HRFlowable(width="100%", thickness=0.5, color=colors.grey))
    story.append(Paragraph(
        f"Relatório gerado automaticamente pelo sistema ARTESP WEB em {agora}. "
        "Dados extraídos do PDF original sem alteração.",
        est["rodape"]
    ))

    doc.build(story)
    return buf.getvalue()


# ─────────────────────────────────────────────────────────────────────────────
# PONTO DE ENTRADA PARA O ROUTER
# ─────────────────────────────────────────────────────────────────────────────

def _montar_resumo_serializavel(ncs: list[NcItem],
                                alertas_km: list[GapAlerta],
                                alertas_codigo: list[CodigoGapAlerta]) -> dict:
    """Constrói resumo JSON-serializável."""
    res = resumo_estatistico(ncs)
    res["alertas_gap"] = [
        {
            "grupo":     a.grupo,
            "empresa":   a.empresa,
            "rodovia":   a.rodovia,
            "sentido":   a.sentido,
            "km_antes":  _km_fmt(a.km_antes),
            "km_depois": _km_fmt(a.km_depois),
            "gap_km":    a.gap_km,
        }
        for a in alertas_km
    ]
    res["alertas_codigo"] = [
        {
            "grupo":            a.grupo,
            "empresa":          a.empresa,
            "codigo_antes":     a.codigo_antes,
            "codigo_depois":    a.codigo_depois,
            "n_faltantes":      a.n_faltantes,
            "codigos_faltantes": a.codigos_faltantes,
        }
        for a in alertas_codigo
    ]
    res["total_ocultos"] = sum(a.n_faltantes for a in alertas_codigo)
    res["emergenciais_lista"] = [
        {
            "codigo":     nc.codigo,
            "rodovia":    nc.rodovia,
            "km":         _km_fmt(nc.km_ini),
            "atividade":  nc.atividade,
            "prazo":      nc.prazo_str,
            "prazo_dias": nc.prazo_dias,
        }
        for nc in res.get("emergenciais", [])
    ]
    res.pop("emergenciais", None)
    res.pop("panelas", None)
    return res


def _excel_valor_para_horario(val) -> str:
    """Converte valor da célula Excel (serial, datetime.time/datetime ou string) para HH:MM ou HH:MM:SS."""
    if val is None:
        return ""
    try:
        from datetime import time, datetime
        if isinstance(val, time):
            return val.strftime("%H:%M:%S") if val.second else val.strftime("%H:%M")
        if isinstance(val, datetime):
            return val.strftime("%H:%M:%S") if val.second else val.strftime("%H:%M")
    except Exception:
        pass
    if isinstance(val, (int, float)):
        v = float(val)
        if v >= 1:
            v = v % 1
        if 0 <= v < 1:
            h = int(v * 24) % 24
            m = int((v * 24 * 60) % 60)
            s = int((v * 24 * 3600) % 60)
            if s:
                return f"{h:02d}:{m:02d}:{s:02d}"
            return f"{h:02d}:{m:02d}"
        if 0 <= v < 24 and v == int(v):
            return f"{int(v):02d}:00"
    s = str(val).strip()
    if re.match(r"^\d{1,2}:\d{2}", s):
        return s
    return s


# Template: lista de campos exigidos; leitura (Excel/PDF) mapeia colunas/blocos para esses campos.
CAMPOS_TEMPLATE_LIST: tuple[str, ...] = (
    "codigo", "data_con", "horario_fiscalizacao", "rodovia", "concessionaria",
    "km_ini_str", "km_fim_str", "sentido", "tipo_atividade", "grupo_atividade",
    "atividade", "prazo_str", "empresa", "nome_fiscal",
)

# (campo, termos no cabeçalho para detectar coluna, termos que desqualificam). Primeiro match ganha.
MAPEAMENTO_EXCEL_PARA_TEMPLATE: list[tuple[str, list[str], list[str]]] = [
    ("codigo", ["cod", "fiscal", "codigo"], []),
    ("codigo", ["cod", "fiscal"], []),
    ("data_con", ["data", "constata"], []),
    ("data_con", ["data", "fiscaliz"], []),
    ("horario_fiscalizacao", ["horario", "hora"], []),
    ("rodovia", ["rodovia", "sp "], []),
    ("rodovia", ["rodovia"], []),
    ("concessionaria", ["concessionaria", "concessionária", "lote"], []),
    ("concessionaria", ["concessionaria", "lote"], []),
    ("km_ini_str", ["km", "inicial", "km+m"], []),
    ("km_ini_str", ["inicial"], []),
    ("km_fim_str", ["km", "final", "km+m"], []),
    ("km_fim_str", ["final"], []),
    ("sentido", ["sentido"], []),
    ("tipo_atividade", ["tipo", "atividade"], []),
    ("grupo_atividade", ["grupo", "atividade"], []),
    ("atividade", ["atividade"], ["tipo", "grupo"]),
    ("atividade", ["evento"], []),
    ("prazo_str", ["prazo", "data limite", "data programada", "termino"], []),
    ("prazo_str", ["data", "termino"], []),
    ("empresa", ["empresa", "fiscalizadora"], []),
    ("nome_fiscal", ["responsavel", "responsável", "tecnico"], []),
    ("nome_fiscal", ["responsavel", "tecnico"], []),
]


def _cel_val(ws_or_sh, is_xlrd: bool, row: int, col: int):
    if is_xlrd:
        if row - 1 >= getattr(ws_or_sh, "nrows", 0) or col - 1 >= getattr(ws_or_sh, "ncols", 0):
            return None
        return ws_or_sh.cell_value(row - 1, col - 1)
    return ws_or_sh.cell(row=row, column=col).value


def _detectar_colunas_template_excel(ws_or_sh, is_xlrd: bool, max_col: int = 30) -> dict[str, Optional[int]]:
    """Cabeçalho (linhas 1–4) → número da coluna por campo do template. None se não encontrado."""
    out: dict[str, Optional[int]] = {attr: None for attr in CAMPOS_TEMPLATE_LIST}
    for row_idx in (2, 3, 1, 4):
        for c in range(1, max_col + 1):
            val = _cel_val(ws_or_sh, is_xlrd, row_idx, c)
            s = (str(val or "")).strip().lower()
            for char, repl in (("é", "e"), ("á", "a"), ("ó", "o"), ("í", "i"), ("ú", "u"), ("ã", "a"), ("õ", "o"), ("ç", "c")):
                s = s.replace(char, repl)
            if not s:
                continue
            for attr, termos, exclude in MAPEAMENTO_EXCEL_PARA_TEMPLATE:
                if out[attr] is not None:
                    continue
                if exclude and any(ex in s for ex in exclude):
                    continue
                if all(t in s for t in termos) or (len(termos) == 1 and termos[0] in s):
                    out[attr] = c
                    break
    return out


def _detectar_colunas_cabecalho(ws_or_sh, is_xlrd: bool, max_col: int = 25) -> tuple[Optional[int], Optional[int], Optional[int], Optional[int]]:
    """
    Detecta índices das colunas Código (C), Horário (E), Tipo (O), Grupo (P) pelo cabeçalho (linhas 1–4).
    Retorna None para colunas não encontradas — planilhas com menos colunas preenchem só o que existir no template.
    """
    cols = _detectar_colunas_template_excel(ws_or_sh, is_xlrd, max_col=max_col)
    return (
        cols.get("codigo"),
        cols.get("horario_fiscalizacao"),
        cols.get("tipo_atividade"),
        cols.get("grupo_atividade"),
    )


def _colunas_disponiveis_no_arquivo(col_map: dict[str, Optional[int]], ncols: int) -> set[str]:
    """Campos do template com coluna identificada neste arquivo."""
    return {
        attr for attr in CAMPOS_TEMPLATE_LIST
        if col_map.get(attr) is not None and col_map[attr] <= ncols
    }


def _ler_excel_complementar(excel_bytes: bytes) -> list[dict]:
    """Lê linhas do Excel; colunas identificadas pelo cabeçalho (MAPEAMENTO_EXCEL_PARA_TEMPLATE). Chaves = CAMPOS_TEMPLATE_LIST."""
    PRIMEIRA_LINHA_DADOS = 5
    out: list[dict] = []
    if not excel_bytes or len(excel_bytes) < 100:
        return out
    try:
        buf = io.BytesIO(excel_bytes)
        is_xls = len(excel_bytes) >= 8 and excel_bytes[:8] == b'\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1'
        if is_xls:
            import xlrd
            rb = xlrd.open_workbook(file_contents=excel_bytes)
            sh = rb.sheet_by_index(0)
            ncols = getattr(sh, "ncols", 0)
            col_map = _detectar_colunas_template_excel(sh, is_xlrd=True, max_col=min(30, ncols))
            disponiveis = _colunas_disponiveis_no_arquivo(col_map, ncols)
            col_c = col_map.get("codigo") or 1
            if col_c - 1 >= ncols:
                return out
            for r in range(PRIMEIRA_LINHA_DADOS - 1, sh.nrows):
                cod = sh.cell_value(r, col_c - 1)
                if cod is None or (isinstance(cod, str) and not str(cod).strip()):
                    continue
                row_dict: dict[str, str] = {attr: "" for attr in CAMPOS_TEMPLATE_LIST}
                row_dict["codigo"] = str(cod).strip()
                for attr in disponiveis:
                    if attr == "codigo":
                        continue
                    col = col_map[attr]
                    if col is None or col - 1 >= ncols:
                        continue
                    val = sh.cell_value(r, col - 1)
                    if val is None and attr != "codigo":
                        continue
                    s = str(val or "").strip()
                    if attr == "horario_fiscalizacao":
                        s = _excel_valor_para_horario(val) or s
                    if s:
                        row_dict[attr] = s[:200] if attr in ("tipo_atividade", "atividade") else s[:100] if attr == "grupo_atividade" else s
                ci = col_map.get("km_ini_str")
                if ci and ci <= ncols:
                    try:
                        v1 = sh.cell_value(r, ci - 1)
                        v2 = sh.cell_value(r, ci) if ci < ncols else None
                        if v1 is not None and v2 is not None:
                            row_dict["km_ini_str"] = "{} + {}".format(int(float(v1)), int(float(v2)))
                    except (ValueError, TypeError):
                        pass
                cf = col_map.get("km_fim_str")
                if cf and cf <= ncols:
                    try:
                        v1 = sh.cell_value(r, cf - 1)
                        v2 = sh.cell_value(r, cf) if cf < ncols else None
                        if v1 is not None and v2 is not None:
                            row_dict["km_fim_str"] = "{} + {}".format(int(float(v1)), int(float(v2)))
                    except (ValueError, TypeError):
                        pass
                out.append(row_dict)
        else:
            if not OPENPYXL_OK:
                return out
            wb = openpyxl.load_workbook(buf, read_only=False, data_only=True)
            ws = wb.active
            max_col_ws = getattr(ws, "max_column", None) or 30
            col_map = _detectar_colunas_template_excel(ws, is_xlrd=False, max_col=min(30, max_col_ws))
            disponiveis = _colunas_disponiveis_no_arquivo(col_map, max_col_ws)
            col_c = col_map.get("codigo") or 1
            max_row = ws.max_row or 0
            for r in range(PRIMEIRA_LINHA_DADOS, max_row + 1):
                cod = ws.cell(row=r, column=col_c).value
                if cod is None or (isinstance(cod, str) and not cod.strip()):
                    continue
                row_dict = {attr: "" for attr in CAMPOS_TEMPLATE_LIST}
                row_dict["codigo"] = str(cod).strip()
                for attr in disponiveis:
                    if attr == "codigo":
                        continue
                    col = col_map.get(attr)
                    if col is None:
                        continue
                    val = ws.cell(row=r, column=col).value
                    if val is None:
                        continue
                    s = str(val or "").strip()
                    if attr == "horario_fiscalizacao":
                        s = _excel_valor_para_horario(val) or s
                    if s:
                        row_dict[attr] = s[:200] if attr in ("tipo_atividade", "atividade") else s[:100] if attr == "grupo_atividade" else s
                ci = col_map.get("km_ini_str")
                if ci and ci < max_col_ws:
                    try:
                        v1 = ws.cell(row=r, column=ci).value
                        v2 = ws.cell(row=r, column=ci + 1).value
                        if v1 is not None and v2 is not None:
                            row_dict["km_ini_str"] = "{} + {}".format(int(float(v1)), int(float(v2)))
                    except (ValueError, TypeError):
                        pass
                cf = col_map.get("km_fim_str")
                if cf and cf < max_col_ws:
                    try:
                        v1 = ws.cell(row=r, column=cf).value
                        v2 = ws.cell(row=r, column=cf + 1).value
                        if v1 is not None and v2 is not None:
                            row_dict["km_fim_str"] = "{} + {}".format(int(float(v1)), int(float(v2)))
                    except (ValueError, TypeError):
                        pass
                out.append(row_dict)
            wb.close()
    except Exception as e:
        logger.warning("Excel complementar não pôde ser lido: %s", e)
    return out


def analisar_e_gerar_pdf(pdf_bytes: bytes,
                          limiar_km: float = LIMIAR_GAP_KM) -> tuple[bytes, bytes, dict]:
    """Pipeline completo para um único PDF. Retorna (pdf_bytes, xlsx_bytes, resumo_dict)."""
    return analisar_e_gerar_pdf_multi([pdf_bytes], limiar_km=limiar_km)


def _norm_codigo_fiscalizacao(c) -> str:
    """Normaliza código da fiscalização para cruzamento PDF–Excel (ex.: 906290.0 → '906290')."""
    if c is None:
        return ""
    s = str(c).strip()
    try:
        f = float(s)
        if f == int(f):
            return str(int(f))
        return s
    except ValueError:
        return s


def _nc_item_desde_excel(row: dict) -> NcItem:
    """Cria NcItem a partir de uma linha do Excel (PDF irmão: código existe só no Excel)."""
    cod = (row.get("codigo") or "").strip()
    km_ini_str = (row.get("km_ini_str") or "").strip()
    km_fim_str = (row.get("km_fim_str") or "").strip()
    km_ini = _km_para_float(km_ini_str)
    km_fim = _km_para_float(km_fim_str)
    nc = NcItem(
        codigo=cod,
        data_con=(row.get("data_con") or "").strip()[:50],
        horario_fiscalizacao=(row.get("horario_fiscalizacao") or "").strip()[:50],
        km_ini_str=km_ini_str[:50],
        km_fim_str=km_fim_str[:50] or km_ini_str[:50],
        km_ini=km_ini,
        km_fim=km_fim if km_fim else km_ini,
        sentido=(row.get("sentido") or "").strip()[:50],
        atividade=(row.get("atividade") or "").strip()[:200],
        tipo_atividade=(row.get("tipo_atividade") or "").strip()[:100],
        grupo_atividade=(row.get("grupo_atividade") or "").strip()[:100],
        observacao="",
        rodovia=(row.get("rodovia") or "").strip()[:50],
        rodovia_nome="",
        lote=(row.get("concessionaria") or "").strip()[:50],
        concessionaria=(row.get("concessionaria") or "").strip()[:50],
        prazo_str=(row.get("prazo_str") or "").strip()[:50],
        prazo_dias=None,
        emergencial=False,
        tipo_panela=False,
        grupo=0,
        empresa="",
        nome_fiscal=(row.get("nome_fiscal") or "").strip()[:100],
        origem_ma=False,
    )
    if (nc.km_ini_str or nc.km_ini) and (not nc.km_fim_str or nc.km_fim == 0.0):
        nc.km_fim_str = nc.km_ini_str or ""
        nc.km_fim = nc.km_ini
    return nc


def _ncs_ma_para_nc_items(ncs_ma: list) -> list[NcItem]:
    """Converte NCs de Meio Ambiente (NcItemMA) para NcItem para o relatório de análise."""
    from .analisar_pdf_ma import NcItemMA
    out: list[NcItem] = []
    for a in ncs_ma:
        if not isinstance(a, NcItemMA):
            continue
        nc = NcItem(
            codigo=(a.codigo_fiscalizacao or a.codigo or "").strip(),
            data_con=a.data_con or "",
            horario_fiscalizacao=getattr(a, "horario_fiscalizacao", "") or "",
            km_ini_str=a.km_ini_str or "",
            km_fim_str=a.km_fim_str or "",
            km_ini=a.km_ini,
            km_fim=a.km_fim,
            sentido=a.sentido or "",
            atividade=a.atividade or "",
            tipo_atividade=getattr(a, "tipo_atividade", "") or "",
            grupo_atividade=getattr(a, "grupo_atividade", "") or "",
            observacao=(a.complemento or a.relatorio or "").strip()[:200],
            rodovia=a.rodovia or "",
            rodovia_nome="",
            lote="",
            concessionaria="",
            prazo_str=a.prazo_str or "",
            prazo_dias=a.prazo_dias,
            emergencial=(a.prazo_dias is not None and a.prazo_dias <= PRAZO_EMERG_MAX),
            tipo_panela=_is_panela(a.atividade or ""),
            grupo=a.grupo,
            empresa=a.empresa or "",
            nome_fiscal=getattr(a, "nome_fiscal", "") or "",
            origem_ma=True,
        )
        if (nc.km_ini_str or nc.km_ini) and (not nc.km_fim_str or nc.km_fim == 0.0):
            nc.km_fim_str = nc.km_ini_str or ""
            nc.km_fim = nc.km_ini
        out.append(nc)
    return out


def analisar_e_gerar_pdf_multi(pdfs_bytes: list[bytes],
                                limiar_km: float = LIMIAR_GAP_KM,
                                nomes: list[str] | None = None,
                                lote: str | None = None,
                                excel_bytes: bytes | list[bytes] | None = None) -> tuple[bytes, bytes, dict]:
    """
    Pipeline completo para múltiplos PDFs.
    Se excel_bytes for informado (Excel que acompanha os PDFs), preenche horário (E), tipo (O) e grupo (P)
    a partir do Excel por correspondência de código fiscalização.
    Retorna (pdf_relatorio_bytes, xlsx_bytes, resumo_dict).
    """
    from .analisar_pdf_ma import parse_pdf_ma
    # Mapa EAF e responsáveis do lote selecionado (13, 21, 26). Lote 13 = config atual; 21/26 quando preenchidos.
    try:
        from nc_artesp.config import get_mapa_eaf, get_mapa_responsavel_tecnico
    except ImportError:
        try:
            from nc_artesp.config import MAPA_RESPONSAVEL_TECNICO as _RT_PADRAO
        except ImportError:
            _RT_PADRAO = {}
        get_mapa_eaf = lambda l: _MAPA_EAF_PADRAO
        get_mapa_responsavel_tecnico = lambda l: _RT_PADRAO
    _lote_num = re.search(r"\d+", (lote or "").strip()) if lote else None
    lote_num = (_lote_num.group(0) if _lote_num else "").strip() or "13"
    mapa_eaf_lote = get_mapa_eaf(lote_num) or _MAPA_EAF_PADRAO
    mapa_responsavel_lote = get_mapa_responsavel_tecnico(lote_num) or {}

    # Limpeza automática: relatório usa somente os PDFs desta requisição.
    ncs_total: list[NcItem] = []
    pdfs_list = list(pdfs_bytes)
    for i, pdf_bytes in enumerate(pdfs_list):
        src = (nomes[i] if nomes and i < len(nomes) else f"PDF {i + 1}")
        parcial = parse_pdf_nc(pdf_bytes)
        if not parcial:
            # Tentar como PDF de Meio Ambiente (grupo já atribuído em parse_pdf_ma)
            parcial_ma = parse_pdf_ma(pdf_bytes)
            if parcial_ma:
                parcial = _ncs_ma_para_nc_items(parcial_ma)
        for nc in parcial:
            setattr(nc, "_origem", src)
        ncs_total.extend(parcial)

    # Atribui EAF/Grupo (para NCs de Conservação; MA já vem com grupo/empresa)
    for nc in ncs_total:
        _atribuir_grupo(nc, mapa_eaf_lote)

    # Ordena: Grupo → Rodovia → Sentido → KM
    ncs_total.sort(key=lambda n: (n.grupo or 999, n.rodovia, n.sentido, n.km_ini))

    vistos: dict[str, NcItem] = {}
    for nc in ncs_total:
        cod = (nc.codigo or "").strip()
        if not cod:
            vistos[f"_idx_{len(vistos)}"] = nc
            continue
        if cod not in vistos:
            vistos[cod] = nc
        else:
            existente = vistos[cod]
            if not existente.horario_fiscalizacao and nc.horario_fiscalizacao:
                existente.horario_fiscalizacao = nc.horario_fiscalizacao
            if not existente.tipo_atividade and nc.tipo_atividade:
                existente.tipo_atividade = nc.tipo_atividade
            if not existente.grupo_atividade and nc.grupo_atividade:
                existente.grupo_atividade = nc.grupo_atividade
            # Trecho (rodovia/km) define EAF; preencher lacunas ou preferir NEP/EBP 22 sobre Autoroutes
            if not (existente.rodovia or "").strip() and (nc.rodovia or "").strip():
                existente.rodovia = nc.rodovia
                if nc.km_ini is not None:
                    existente.km_ini = nc.km_ini
                if getattr(nc, "km_fim", None) is not None:
                    existente.km_fim = nc.km_fim
            elif existente.km_ini is None and nc.km_ini is not None and (nc.rodovia or "").strip():
                existente.rodovia = existente.rodovia or nc.rodovia
                existente.km_ini = nc.km_ini
                if getattr(nc, "km_fim", None) is not None:
                    existente.km_fim = nc.km_fim
            elif (existente.rodovia or "").strip() and (nc.rodovia or "").strip() and existente.km_ini is not None and nc.km_ini is not None:
                from nc_artesp.utils.helpers import obter_grupo_empresa_por_trecho
                _ge, _ee = obter_grupo_empresa_por_trecho(existente.rodovia, existente.km_ini, mapa_eaf_lote)
                _gn, _en = obter_grupo_empresa_por_trecho(nc.rodovia, nc.km_ini, mapa_eaf_lote)
                if _ee == "Autoroutes" and _en in ("NEP", "EBP 22"):
                    existente.rodovia = nc.rodovia
                    existente.km_ini = nc.km_ini
                    if getattr(nc, "km_fim", None) is not None:
                        existente.km_fim = nc.km_fim
    ncs_total = list(vistos.values())
    ncs_total.sort(key=lambda n: (n.grupo or 999, n.rodovia, n.sentido, n.km_ini))

    # Cruzamento por código: uma NC por código; lacunas do PDF preenchidas do Excel. EAF/grupo só por trecho.
    if excel_bytes:
        if isinstance(excel_bytes, bytes):
            excel_bytes = [excel_bytes]
        linhas_excel: list[dict] = []
        for buf in excel_bytes:
            linhas_excel.extend(_ler_excel_complementar(buf))
        if linhas_excel:
            por_codigo_excel = {
                _norm_codigo_fiscalizacao(r.get("codigo")): r
                for r in linhas_excel
                if r.get("codigo") and _norm_codigo_fiscalizacao(r.get("codigo"))
            }
            # Só Excel (sem PDF): construir todas as NCs a partir do Excel
            if not ncs_total:
                for cod_excel, row in por_codigo_excel.items():
                    if cod_excel:
                        nc_excel = _nc_item_desde_excel(row)
                        setattr(nc_excel, "_origem", "Excel")
                        ncs_total.append(nc_excel)
                ncs_total.sort(key=lambda n: (n.grupo or 999, n.rodovia, n.sentido, n.km_ini))
            else:
                codigos_pdf = {_norm_codigo_fiscalizacao(nc.codigo) for nc in ncs_total if (nc.codigo or "").strip()}

                for nc in ncs_total:
                    cod = _norm_codigo_fiscalizacao(nc.codigo)
                    if not cod:
                        continue
                    row = por_codigo_excel.get(cod)
                    if not row:
                        continue
                    for attr in CAMPOS_TEMPLATE_LIST:
                        if attr == "codigo" or attr == "empresa":
                            continue
                        val_nc = (getattr(nc, attr, None) or "") if hasattr(nc, attr) else ""
                        val_excel = (row.get(attr) or "").strip()
                        if not (val_nc or "").strip() and val_excel:
                            if attr in ("tipo_atividade", "atividade"):
                                setattr(nc, attr, val_excel[:200])
                            elif attr == "grupo_atividade":
                                setattr(nc, attr, val_excel[:100])
                            else:
                                setattr(nc, attr, val_excel[:100])
                            if attr in ("km_ini_str", "km_fim_str") and hasattr(nc, "km_ini"):
                                try:
                                    k = _km_para_float(val_excel)
                                    if attr == "km_ini_str":
                                        nc.km_ini = k
                                    else:
                                        nc.km_fim = k
                                except Exception:
                                    pass

                # Códigos que estão só no Excel: criar NC a partir da linha do Excel (PDF irmão)
                for cod_excel, row in por_codigo_excel.items():
                    if cod_excel and cod_excel not in codigos_pdf:
                        nc_excel = _nc_item_desde_excel(row)
                        setattr(nc_excel, "_origem", "Excel")
                        ncs_total.append(nc_excel)
                ncs_total.sort(key=lambda n: (n.grupo or 999, n.rodovia, n.sentido, n.km_ini))

    # EAF/grupo por trecho (rodovia+km). Trava: se EAF_PERMITIDAS estiver preenchida no config,
    # qualquer empresa fora da lista tem grupo e empresa zerados (ex.: só Autoroutes permitida).
    # EAF_PERMITIDAS vazia = todas as EAFs do MAPA_EAF são aceitas.
    try:
        from nc_artesp.config import EAF_PERMITIDAS
        permitidas = set((e or "").strip() for e in EAF_PERMITIDAS if (e or "").strip())
    except Exception:
        permitidas = set()
    for nc in ncs_total:
        if (nc.rodovia or "").strip() and nc.km_ini is not None:
            _atribuir_grupo(nc, mapa_eaf_lote)
        if permitidas and (nc.empresa or "").strip() not in permitidas:
            nc.grupo = 0
            nc.empresa = ""

    for nc in ncs_total:
        if (not (nc.grupo_atividade or "").strip() or not (nc.tipo_atividade or "").strip()) and (nc.atividade or "").strip():
            grupo_inf, tipo_inf = _inferir_grupo_tipo_da_atividade(nc.atividade)
            if grupo_inf and not (nc.grupo_atividade or "").strip():
                nc.grupo_atividade = grupo_inf
            if tipo_inf and not (nc.tipo_atividade or "").strip():
                nc.tipo_atividade = tipo_inf

    # Impede gerar relatório de um lote com PDFs de outro: lote nos PDFs deve bater com o selecionado
    if lote and ncs_total:
        lote_sel = (lote or "").strip()
        lotes_pdf = set()
        for nc in ncs_total:
            num = _lote_num_do_pdf(nc)
            if num:
                lotes_pdf.add(num)
        if lotes_pdf:
            if len(lotes_pdf) > 1:
                raise ValueError(
                    "Os PDFs contêm NCs de mais de um lote ({}). Use apenas PDFs do mesmo lote.".format(", ".join(sorted(lotes_pdf)))
                )
            unico = next(iter(lotes_pdf))
            if unico != lote_sel:
                from nc_artesp.config import _LOTE_CONCESSIONARIA
                nome_pdf = _LOTE_CONCESSIONARIA.get(unico, "Lote " + unico)
                nome_sel = _LOTE_CONCESSIONARIA.get(lote_sel, "Lote " + lote_sel)
                raise ValueError(
                    "Os PDFs são do {} (lote {}), mas você selecionou {} (lote {}). "
                    "Corrija o lote no menu ou use PDFs do lote correto.".format(nome_pdf, unico, nome_sel, lote_sel)
                )

    alertas_km     = analisar_gaps(ncs_total, limiar_km=limiar_km, mapa_eaf=mapa_eaf_lote)
    alertas_codigo = analisar_sequencia_codigos(ncs_total)

    res = _montar_resumo_serializavel(ncs_total, alertas_km, alertas_codigo)
    res["n_arquivos"] = len(pdfs_list)
    # Para API/log: alertas só contam quando relatório é do mesmo dia da constatação
    data_c = res.get("data_con", "")
    data_con_dt = _parse_data(data_c)
    data_con_date = data_con_dt.date() if data_con_dt else None
    data_relatorio = date.today()
    res["relatorio_hoje"] = data_con_date is not None and data_relatorio == data_con_date

    # Validação responsável técnico (mapa do lote): só zera se nome pertencer a outra EAF.
    try:
        def _to_nome_list(v: str) -> list[str]:
            """
            Aceita valor único ou múltiplos nomes separados por ';' ou ','.
            Ex.: "A; B" ou "A, B".
            """
            s = (v or "").strip()
            if not s:
                return []
            parts = [p.strip() for p in re.split(r"[;,]", s) if p and p.strip()]
            return parts or [s]

        # empresa_tag -> lista de nomes permitidos (mapa do lote selecionado)
        empresa_para_nomes: dict[str, list[str]] = {
            (emp or "").strip(): _to_nome_list(nome_map)
            for emp, nome_map in (mapa_responsavel_lote or {}).items()
        }

        def _mapear_emp_para_chave(empresa: str) -> str:
            """
            Normaliza o "tag" de EAF para bater com as chaves do mapa.
            Ex.: "AUTOROUTES G2" -> "Autoroutes"
            """
            e = (empresa or "").strip()
            if not e:
                return ""
            # normaliza espaços/case para facilitar match
            e_norm = re.sub(r"\s+", " ", e)

            # remove sufixos comuns que aparecem em OCR/labels (ex.: "Autoroutes G2")
            e_norm = re.sub(r"\bG\s*\d+\b", "", e_norm, flags=re.IGNORECASE).strip()

            # match exato primeiro (com normalização)
            if e_norm in empresa_para_nomes:
                return e_norm
            if e in empresa_para_nomes:
                return e
            e_up = e_norm.upper()
            for k in empresa_para_nomes.keys():
                kk = (k or "").strip()
                if not kk:
                    continue
                kk_up = kk.upper()
                if kk_up in e_up or e_up in kk_up:
                    return k
            return e_norm  # fallback: não encontrado

        # União (apenas para decidir se o nome existe em algum lugar da regra)
        responsaveis_regra = set()
        for nomes in empresa_para_nomes.values():
            for n in nomes:
                nn = (n or "").strip()
                if nn:
                    responsaveis_regra.add(nn)

        def _nome_na_lista_por_substring(nome: str, lista_nomes: list[str]) -> bool:
            n = (nome or "").strip()
            if not n:
                return False
            for r in (lista_nomes or []):
                rr = (r or "").strip()
                if not rr:
                    continue
                if n == rr:
                    return True
                if n in rr or rr in n:
                    return True
            return False

        for nc in ncs_total:
            emp_raw = (nc.empresa or "").strip()
            emp = _mapear_emp_para_chave(emp_raw)
            if not (nc.nome_fiscal or "").strip():
                nc.nome_fiscal = (mapa_responsavel_lote.get(emp) or "").strip()

            nome = (nc.nome_fiscal or "").strip()
            if not nome:
                continue

            nomes_ok = empresa_para_nomes.get(emp, [])
            if _nome_na_lista_por_substring(nome, nomes_ok):
                continue
            # Nome não está na lista desta EAF. Só zera se o nome pertencer a OUTRA EAF (conflito).
            # Se não estiver em nenhum mapa, mantém empresa do trecho (não zera por nome desconhecido).
            nome_em_outra = False
            for outra_emp, lista_outra in empresa_para_nomes.items():
                if outra_emp == emp or not lista_outra:
                    continue
                if _nome_na_lista_por_substring(nome, lista_outra):
                    nome_em_outra = True
                    break
            if nome_em_outra:
                nc.grupo = 0
                nc.empresa = ""
    except Exception:
        pass
    pdf_rel = gerar_relatorio_pdf(ncs_total, alertas_km, alertas_codigo,
                                  limiar_km=limiar_km, mapa_eaf=mapa_eaf_lote)
    lote_ok = (lote or "").strip() or None
    xlsx_bytes = gerar_relatorio_xlsx(ncs_total, lote_selecionado=lote_ok)
    return pdf_rel, xlsx_bytes, res
