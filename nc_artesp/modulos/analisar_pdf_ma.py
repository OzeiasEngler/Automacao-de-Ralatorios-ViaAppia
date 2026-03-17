"""
nc_artesp/modulos/analisar_pdf_ma.py
────────────────────────────────────────────────────────────────────────────
Extração de INFORMAÇÕES EM TEXTO do PDF de Meio Ambiente (não é extração de imagem;
imagens são tratadas no fluxo Extrair PDF). Coleta alinhada à macro M07 (Kria2_Inserir_NC_MA).
Lógica: PDF inteiro em uma string, parse por regex. Retorna registros (código, rodovia, km,
data, atividade, etc.). Usado por inserir_nc_kria para preencher a planilha passo 1 (EAF) e Kcor-Kria.
"""

from __future__ import annotations

import re
from dataclasses import dataclass
from typing import Optional

try:
    import fitz
    FITZ_OK = True
except ImportError:
    FITZ_OK = False

try:
    import pdfplumber
    PDFPLUMBER_OK = True
except ImportError:
    PDFPLUMBER_OK = False
    pdfplumber = None

try:
    from nc_artesp.config import MAPA_EAF as _MAPA_EAF_MA
except ImportError:
    try:
        from config import MAPA_EAF as _MAPA_EAF_MA
    except ImportError:
        _MAPA_EAF_MA = []


@dataclass
class NcItemMA:
    codigo: str = ""  # Código da NC (ex.: NC.02.2024)
    codigo_fiscalizacao: str = ""  # Código da Fiscalização (extraído do PDF, distinto da NC)
    rodovia: str = ""
    km_ini_str: str = ""
    km_ini: float = 0.0
    km_fim_str: str = ""
    km_fim: float = 0.0
    sentido: str = ""
    data_con: str = ""
    atividade: str = ""
    relatorio: str = ""
    complemento: str = ""
    embasamento: str = ""
    prazo_str: str = ""
    prazo_dias: Optional[int] = None
    # Grupo EAF (fiscalização) — preenchido por _atribuir_grupo_ma para relatório de análise
    grupo: int = 0
    empresa: str = ""
    nome_fiscal: str = ""  # nome do fiscal / responsável técnico (extraído do PDF quando possível)


# Mapeamento sentido (letra → nome completo) — usado na extração e ao montar registro para gravação
SENTIDO_PARA_TEXTO = {"L": "Leste", "O": "Oeste", "N": "Norte", "S": "Sul", "I": "Interna", "E": "Externa"}


def _sentido_para_texto(s: str) -> str:
    """Converte letra (L/O/N/S/I/E) para nome completo. Se já for nome ou vazio, devolve como está."""
    s = (s or "").strip().upper()
    if not s:
        return ""
    letra = s[0] if s else ""
    if letra == "0":
        letra = "O"
    return SENTIDO_PARA_TEXTO.get(letra, s)


def _atribuir_grupo_ma(ncs: list["NcItemMA"], mapa_eaf: list[dict]) -> None:
    """
    Atribui grupo EAF e empresa a cada NC de Meio Ambiente com base em rodovia + km.
    Usa nc_artesp.utils.helpers.obter_grupo_empresa_por_trecho (mesmo MAPA_EAF, Contatos EAFs).
    """
    from nc_artesp.utils.helpers import obter_grupo_empresa_por_trecho
    if not mapa_eaf or not ncs:
        return
    for nc in ncs:
        nc.grupo, nc.empresa = obter_grupo_empresa_por_trecho(nc.rodovia, nc.km_ini, mapa_eaf)


def _km_para_float(s: str) -> float:
    try:
        s = (s or "").strip().replace(" ", "")
        m = re.match(r"(\d+)\+?(\d*)", s)
        if m:
            km = float(m.group(1) or 0)
            met = m.group(2)
            if met:
                met = met.ljust(3, "0")[:3]
                km += float(met) / 1000.0
            return km
    except (ValueError, TypeError):
        pass
    return 0.0


# ─── Extração de texto (fallback / debug) ────────────────────────────────────

def extrair_texto_pdf(pdf_bytes: bytes) -> str:
    if not FITZ_OK:
        return ""
    try:
        doc = fitz.open(stream=pdf_bytes, filetype="pdf")
        partes = [p.get_text("text") for p in doc]
        doc.close()
        return "\n".join(partes or [])
    except Exception:
        return ""


def extrair_texto_pdf_por_blocos(pdf_bytes: bytes) -> str:
    if not FITZ_OK:
        return ""
    try:
        doc = fitz.open(stream=pdf_bytes, filetype="pdf")
        partes = []
        for pag in doc:
            for b in pag.get_text("blocks") or []:
                if len(b) >= 5 and (b[4] or "").strip():
                    partes.append((b[4] or "").strip().replace("\n", " "))
        doc.close()
        return "\n".join(partes)
    except Exception:
        return ""


def extrair_texto_pdf_pdfplumber(pdf_bytes: bytes) -> str:
    if not PDFPLUMBER_OK or not pdfplumber:
        return ""
    import io
    try:
        with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
            return "\n".join(p.extract_text() or "" for p in pdf.pages)
    except Exception:
        return ""


# ─── Coleta principal (1 PDF = 1 registro) ────────────────────────────────────

def extrair_dados_ma(pdf_bytes: bytes) -> Optional[dict]:
    """
    Extrai dados de um PDF da ARTESP (Meio Ambiente).
    PDF inteiro em uma string, primeira NC, regex de pulo para vírgulas.
    Retorna 1 registro por PDF ou None.
    """
    if not PDFPLUMBER_OK or not pdfplumber:
        return None
    import io
    try:
        with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
            texto_completo = " ".join(p.extract_text() or "" for p in pdf.pages)
    except Exception:
        return None
    texto_limpo = " ".join(texto_completo.split()).replace('"', '')

    match_nc = re.search(r"(?:NC|SS|MA|HE|RE)\.\d{2}\.\d{4}", texto_limpo)
    if not match_nc:
        return None
    codigo_nc = match_nc.group(0)

    # Âncoras mais fortes para não perder informação (extrair_dados_estritos_v2)
    rodovia_match = re.search(r"Rodovia\s*\(SP\)\s*:\s*(SP\s*\d+)", texto_limpo, re.I)
    if not rodovia_match:
        rodovia_match = re.search(r"Rodovia.*?\s(SP\s*\d+)", texto_limpo, re.I)
    rod_val = (rodovia_match.group(1) or "").strip() if rodovia_match else ""
    rodovia = rod_val.replace(" ", "-") if rod_val else ""

    # Km + Sentido: L/O/N/S/I/E ou "0" (OCR confunde O com zero) → normalizar e expandir para texto
    loc_match = re.search(r"Km\s*(\d+[\+\d\s]*)\s+([LONSIE0])", texto_limpo)
    km_val = loc_match.group(1).strip() if loc_match else ""
    sentido_letra = (loc_match.group(2) or "").upper().strip() if loc_match else ""
    sentido_val = _sentido_para_texto(sentido_letra)

    # Coletar todas as datas diferentes do PDF; mesma data só uma vez
    datas_encontradas = re.findall(r"\d{2}/\d{2}/\d{4}", texto_limpo)
    datas_unicas = list(dict.fromkeys(datas_encontradas))  # ordem preservada, sem repetir
    data_val = " ; ".join(datas_unicas) if datas_unicas else ""

    patologia_match = re.search(r"Patologia\s*:\s*(.*?)(?=Código da Fiscalização|Código Fiscalização|NR-|FICHA)", texto_limpo, re.S | re.I)
    # Código da Fiscalização (ex.: 902531) — identifica as fotos; não confundir com Num. da NC (HE.13.0112)
    cod_fisc_match = re.search(r"Código\s+da\s+Fiscaliza[cç][aã]o\s*:\s*(\S+)", texto_limpo, re.I)
    if not cod_fisc_match:
        cod_fisc_match = re.search(r"Código\s+Fiscaliza[cç][aã]o\s*:\s*(\S+)", texto_limpo, re.I)
    codigo_fiscalizacao_val = (cod_fisc_match.group(1) or "").strip() if cod_fisc_match else ""
    # Lote nunca é usado como código; só o número após "Lote: " (código da fiscalização)
    if codigo_fiscalizacao_val.upper().startswith("LOTE"):
        lote_match = re.search(r"Lote\s*:\s*(\S+)", texto_limpo, re.I)
        codigo_fiscalizacao_val = (lote_match.group(1) or "").strip() if lote_match else ""
    if not codigo_fiscalizacao_val and FITZ_OK:
        try:
            doc = fitz.open(stream=pdf_bytes, filetype="pdf")
            if len(doc) > 0:
                from nc_artesp.pdf_extractor import _extrair_codigo_nc
                codigo_fiscalizacao_val = (_extrair_codigo_nc(doc[0], doc[0].rect) or "").strip()
            doc.close()
        except Exception:
            pass

    nome_fiscal_val = ""
    for pat in (
        r"Respons[aá]vel\s+T[eé]cnico\s*:\s*([^\n\r|;]+)",
        r"Fiscal\s*:\s*([^\n\r|;]+)",
        r"Respons[aá]vel\s*:\s*([^\n\r|;]+)",
    ):
        m = re.search(pat, texto_limpo, re.I)
        if m:
            nome_fiscal_val = (m.group(1) or "").strip()[:200]
            break

    # Observações: aceitar "Observações" ou "Observação"; capturar até próxima data ou fim (igual macro VBA)
    obs_match = re.search(r"Observa[cç][oõ]es?\s*:\s*(.*?)(?=\d{2}/\d{2}/\d{4}|\Z)", texto_limpo, re.S | re.I)
    pat_texto = (patologia_match.group(1) or "").strip() if patologia_match else ""
    obs_texto = (obs_match.group(1) or "").strip() if obs_match else ""
    # Col U: concatenação como na macro — texto(x) & vbCrLf & "Observações: " & obs(x)
    patologia_completa = f"{pat_texto} | Observações: {obs_texto}".strip() if obs_texto else pat_texto

    return {
        "codigo": codigo_nc,
        "codigo_fiscalizacao": codigo_fiscalizacao_val,
        "nome_fiscal": nome_fiscal_val,
        "rodovia": rodovia,
        "kmi": km_val,
        "sentido": sentido_val,
        "data": data_val,
        "patologia": patologia_completa[:500],
    }


def _dict_para_nc_item(d: dict) -> NcItemMA:
    nc = NcItemMA()
    nc.codigo = (d.get("codigo") or "").strip()
    nc.codigo_fiscalizacao = (d.get("codigo_fiscalizacao") or "").strip()
    nc.nome_fiscal = (d.get("nome_fiscal") or "").strip()[:200]
    nc.rodovia = (d.get("rodovia") or "").strip()
    nc.km_ini_str = (d.get("kmi") or "").strip()
    if nc.km_ini_str:
        nc.km_ini = _km_para_float(nc.km_ini_str)
    nc.sentido = (d.get("sentido") or "").strip()[:30]
    nc.data_con = (d.get("data") or "").strip()
    # Limite 500 chars; macro usa texto + Observações na coluna U
    nc.atividade = (d.get("patologia") or "").strip()[:500]
    return nc


def parse_pdf_ma(pdf_bytes: bytes) -> list[NcItemMA]:
    """Extrai NCs do PDF. 1 PDF = 1 NC (primeira ocorrência). Atribui grupo/empresa via MAPA_EAF."""
    dados = extrair_dados_ma(pdf_bytes)
    if not dados:
        return []
    nc = _dict_para_nc_item(dados)
    if nc.codigo or nc.rodovia or nc.atividade or nc.km_ini_str or nc.data_con:
        _atribuir_grupo_ma([nc], _MAPA_EAF_MA)
        return [nc]
    return []


def parse_pdf_ma_para_registros(pdf_bytes: bytes) -> list[dict]:
    """
    Retorna lista de dicionários para inserir_nc_kria (Kcor-Kria M07).
    1 NC = 1 linha no Excel.
    """
    from nc_artesp.utils.helpers import km_formato_arquivo, formatar_numero

    def _normalizar_rodovia(rod_raw: str) -> tuple[str, str, int]:
        rod = (rod_raw or "").strip()
        mapa = {
            "SP-075": ("SP075", "SP075", 1),
            "SP-127": ("SP127", "SP127", 2),
            "SP-280": ("SP280", "SP280", 3),
            "SP-300": ("SP300", "SP300", 4),
            "SPI-102/300": ("SPI102_300", "SPI102/300", 5),
            "CP-127_147": ("CP-127_147", "FORA", 6),
            "CP-127_308": ("CP-127_308", "FORA", 7),
        }
        rod_norm = re.sub(r"\s+", "-", rod)
        if rod_norm in mapa:
            return mapa[rod_norm]
        rod_norm = re.sub(r"^SP\s*", "SP-", rod, flags=re.IGNORECASE)
        if rod_norm in mapa:
            return mapa[rod_norm]
        tag = re.sub(r"[\s\-/]+", "", rod).replace("/", "_")
        return (tag, rod, 0)

    def _parse_data(s: str):
        from datetime import datetime
        if not s:
            return None
        try:
            return datetime.strptime(s.strip(), "%d/%m/%Y")
        except (ValueError, TypeError):
            return None

    def _primeira_data(s: str) -> str:
        """De 'dd/mm/aaaa ; dd/mm/aaaa' retorna só a primeira data para parsing."""
        if not s:
            return ""
        return (s.split(";")[0] or "").strip() or s.strip()

    ncs = parse_pdf_ma(pdf_bytes)
    registros = []
    for idx, nc in enumerate(ncs, start=1):
        rod_tag, rod_cod, n = _normalizar_rodovia(nc.rodovia)
        dt = _parse_data(_primeira_data(nc.data_con))
        data_str = nc.data_con or ""
        kminicial_arq = km_formato_arquivo(nc.km_ini_str) if nc.km_ini_str else ""
        numero = formatar_numero(nc.codigo or idx, 6)
        atividade_limpa = " ".join((nc.atividade or "").split())
        resumo = atividade_limpa[:150].strip() if atividade_limpa else ""

        # Garantir sentido gravado como texto (Leste/Oeste/Norte/Sul/Interna/Externa), não letra
        sentido_gravacao = _sentido_para_texto(nc.sentido or "")

        registros.append({
            "y": 0,
            "relatorio": nc.relatorio or "",
            "codigo": nc.codigo or "",
            "codigo_fiscalizacao": nc.codigo_fiscalizacao or "",
            "nome_fiscal": getattr(nc, "nome_fiscal", "") or "",
            "resumo": resumo,
            "complemento": nc.complemento or "",
            "embasamento": nc.embasamento or nc.prazo_str or "",
            "rod_raw": nc.rodovia or "",
            "rod_tag": rod_tag,
            "rod_cod": rod_cod,
            "n": n,
            "texto": atividade_limpa,
            "sentido": sentido_gravacao,
            "kminicial_t": nc.km_ini_str or "",
            "kmfinal_t": nc.km_fim_str or nc.km_ini_str or "",
            "kminicial_arq": kminicial_arq,
            "data_raw": data_str,
            "dt": dt,
            "prazo": nc.prazo_dias if nc.prazo_dias is not None else nc.prazo_str,
            "numero": numero,
            "foto_id": None,
            "serv_nc": "Reclassificar",
            "classifica": "Conservação Rotina",
            "executor": "",
            "grupo": getattr(nc, "grupo", 0),
            "empresa": getattr(nc, "empresa", ""),
        })
    return registros


def ncs_ma_para_dict_m2(ncs: list[NcItemMA]) -> list[dict]:
    """Converte NcItemMA para o formato do gerar_modelo_foto (Kria/Resposta)."""
    from datetime import datetime
    from nc_artesp.utils.helpers import formatar_numero

    def _parse_dt(s: str):
        if not s:
            return None
        try:
            return datetime.strptime((s or "").strip(), "%d/%m/%Y")
        except (ValueError, TypeError):
            return None

    def _primeira_data(s: str) -> str:
        if not s:
            return ""
        return (s.split(";")[0] or "").strip() or s.strip()

    resultado = []
    for idx, nc in enumerate(ncs, start=1):
        rod = (nc.rodovia or "").replace(" ", "-")
        if rod and not rod.startswith("SP-"):
            rod = re.sub(r"^SP\s*", "SP-", rod, flags=re.IGNORECASE)
        num_foto = formatar_numero(nc.codigo or idx, 6)
        resultado.append({
            "codigo": nc.codigo or "",
            "data_con": _parse_dt(_primeira_data(nc.data_con)),
            "data_reparo": _parse_dt(nc.embasamento or nc.prazo_str),
            "tipo_nc": "Conservação Rotina",
            "rod_codigo": rod or "",
            "rod_tag": rod or "",
            "sentido": nc.sentido or "",
            "km_i": nc.km_ini_str or "",
            "km_f": nc.km_fim_str or nc.km_ini_str or "",
            "num_foto": num_foto,
            "prazo_dias": nc.prazo_dias,
        })
    return resultado
