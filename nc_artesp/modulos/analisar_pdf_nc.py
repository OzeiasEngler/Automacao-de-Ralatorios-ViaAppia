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
import re
from dataclasses import dataclass, field
from datetime import datetime, date
from typing import Optional

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
    km_ini_str: str       = ""   # "50 + 950"
    km_fim_str: str       = ""
    km_ini: float         = 0.0  # valor decimal (50.950)
    km_fim: float         = 0.0
    sentido: str          = ""
    atividade: str        = ""
    observacao: str       = ""
    rodovia: str          = ""   # "SP 075"
    rodovia_nome: str     = ""
    lote: str             = ""
    concessionaria: str   = ""
    prazo_str: str        = ""   # DD/MM/AAAA
    prazo_dias: Optional[int] = None
    emergencial: bool     = False
    tipo_panela: bool     = False  # atividade sugere panela/buraco
    # Grupo EAF (coluna V do template EAF — número da fiscalizadora)
    grupo: int            = 0    # 0 = não identificado
    empresa: str          = ""   # nome da empresa fiscalizadora
    origem_ma: bool       = False  # True = PDF Meio Ambiente (não entra em gap KM panelas/buracos)


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


def _extrair_texto_pdf(pdf_bytes: bytes) -> str:
    """Extrai todo o texto do PDF em uma string única."""
    if not FITZ_OK:
        raise ImportError("PyMuPDF não instalado: pip install pymupdf")
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    partes = []
    for pag in doc:
        partes.append(pag.get_text("text"))
    doc.close()
    return "\n".join(partes)


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
    if not nc.sentido and nc.km_ini_str:
        m = re.search(r'Sentido\s*:?\s*([LONSIE0])\b', texto_flat, re.IGNORECASE)
        if m:
            nc.sentido = _sentido_para_texto(m.group(1).strip())
        if not nc.sentido:
            m = re.search(r'\b([LONSIE0])\s+Sentido\b', texto_flat, re.IGNORECASE)
            if m:
                nc.sentido = _sentido_para_texto(m.group(1).strip())

    # ── lote (linha isolada com apenas dígitos, após "Data Limite para Reparo")
    limite_idx = next(
        (i for i, ln in enumerate(lines) if "Limite para Reparo" in ln), -1
    )
    if limite_idx >= 0:
        for ln in lines[limite_idx + 1: limite_idx + 4]:
            if re.match(r'^\d+$', ln):
                nc.lote = ln
                break

    # ── atividade
    for ln in lines:
        m = re.match(r'Atividade:\s*(.+)', ln, re.IGNORECASE)
        if m:
            nc.atividade = m.group(1).strip()
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

    # ── rodovia SP (ex: "SP 075")
    for ln in lines:
        m = re.search(r'Rodovia\s*\(SP\):\s*(.+)', ln, re.IGNORECASE)
        if m:
            nc.rodovia = m.group(1).strip()
            break

    # ── prazo: data isolada APÓS a rodovia SP
    rodsp_idx = next(
        (i for i, ln in enumerate(lines) if "Rodovia (SP):" in ln or
         re.search(r'Rodovia\s*\(SP\)', ln)), -1
    )
    if rodsp_idx >= 0:
        for ln in lines[rodsp_idx + 1: rodsp_idx + 5]:
            if re.match(r'^\d{2}/\d{2}/\d{4}$', ln):
                nc.prazo_str = ln
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

    # NC válida precisa ao menos de código ou atividade
    if not (nc.codigo or nc.atividade):
        return None
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
    """
    Extrai todas as NCs do PDF de Constatação de Rotina Artesp.
    Retorna lista de NcItem ordenada por (rodovia, sentido, km_ini).
    """
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


def _banner(texto: str, cor: colors.Color, estilos: dict):
    """Faixa colorida com texto branco — usada como cabeçalho de seção."""
    return Table(
        [[Paragraph(texto, estilos["secao"])]],
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


def gerar_relatorio_pdf(ncs: list[NcItem],
                        alertas_km: list[GapAlerta],
                        alertas_codigo: list[CodigoGapAlerta],
                        limiar_km: float = LIMIAR_GAP_KM) -> bytes:
    """
    Gera o PDF de análise em bytes.
    Estrutura: Capa → Alertas de KM → NCs emergenciais → NCs por tipo (em seq. de KM).
    Os dados das NCs são reproduzidos sem qualquer alteração.
    """
    if not REPORTLAB_OK:
        raise ImportError("reportlab não instalado: pip install reportlab")

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
        return Paragraph((s or "").replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;"), style)
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
        for entry in _MAPA_EAF_PADRAO:
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
            for entry in _MAPA_EAF_PADRAO:
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


def analisar_e_gerar_pdf(pdf_bytes: bytes,
                          limiar_km: float = LIMIAR_GAP_KM) -> tuple[bytes, dict]:
    """
    Pipeline completo para um único PDF: parse → análise → relatório PDF.
    Retorna (pdf_relatorio_bytes, resumo_dict).
    """
    return analisar_e_gerar_pdf_multi([pdf_bytes], limiar_km=limiar_km)


def _ncs_ma_para_nc_items(ncs_ma: list) -> list[NcItem]:
    """Converte NCs de Meio Ambiente (NcItemMA) para NcItem para o relatório de análise."""
    from .analisar_pdf_ma import NcItemMA
    out: list[NcItem] = []
    for a in ncs_ma:
        if not isinstance(a, NcItemMA):
            continue
        nc = NcItem(
            codigo=a.codigo or "",
            data_con=a.data_con or "",
            km_ini_str=a.km_ini_str or "",
            km_fim_str=a.km_fim_str or "",
            km_ini=a.km_ini,
            km_fim=a.km_fim,
            sentido=a.sentido or "",
            atividade=a.atividade or "",
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
            origem_ma=True,  # Meio Ambiente não aponta panelas/buracos — não entra em gap KM
        )
        if (nc.km_ini_str or nc.km_ini) and (not nc.km_fim_str or nc.km_fim == 0.0):
            nc.km_fim_str = nc.km_ini_str or ""
            nc.km_fim = nc.km_ini
        out.append(nc)
    return out


def analisar_e_gerar_pdf_multi(pdfs_bytes: list[bytes],
                                limiar_km: float = LIMIAR_GAP_KM,
                                nomes: list[str] | None = None) -> tuple[bytes, dict]:
    """
    Pipeline completo para múltiplos PDFs.
    Combina as NCs de todos os arquivos antes de analisar e gerar o relatório.
    Se o PDF for de Conservação (Constatação Rotina), usa parse_pdf_nc.
    Se não extrair NCs, tenta parse_pdf_ma (Meio Ambiente) e converte para o mesmo relatório.
    Análises realizadas:
      1. Atribuição de Grupo EAF por rodovia+km
      2. Gaps de KM por grupo/rodovia/sentido
      3. Gaps de numeração do Código Fiscalização por grupo (apontamentos não entregues)
    Retorna (pdf_relatorio_bytes, resumo_dict).
    """
    from .analisar_pdf_ma import parse_pdf_ma
    # Limpeza automática: relatório usa somente os PDFs desta requisição.
    # Nenhum dado de execuções anteriores é reutilizado (sem cache nem estado persistido).
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
        _atribuir_grupo(nc, _MAPA_EAF_PADRAO)

    # Ordena: Grupo → Rodovia → Sentido → KM
    ncs_total.sort(key=lambda n: (n.grupo or 999, n.rodovia, n.sentido, n.km_ini))

    alertas_km     = analisar_gaps(ncs_total, limiar_km=limiar_km)
    alertas_codigo = analisar_sequencia_codigos(ncs_total)

    res = _montar_resumo_serializavel(ncs_total, alertas_km, alertas_codigo)
    res["n_arquivos"] = len(pdfs_list)
    # Para API/log: alertas só contam quando relatório é do mesmo dia da constatação
    data_c = res.get("data_con", "")
    data_con_dt = _parse_data(data_c)
    data_con_date = data_con_dt.date() if data_con_dt else None
    data_relatorio = date.today()
    res["relatorio_hoje"] = data_con_date is not None and data_relatorio == data_con_date

    pdf_rel = gerar_relatorio_pdf(ncs_total, alertas_km, alertas_codigo,
                                  limiar_km=limiar_km)
    return pdf_rel, res
