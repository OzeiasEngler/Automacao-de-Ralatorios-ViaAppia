"""
Extração do PDF de NC Constatação: Texto (COD).jpg, PDF (COD).jpg, nc (COD).jpg.
Dimensões/DPI finais = M02_FOTO_* / M02_FOTO_PDF_* (ARTESP e Artemig lote 50).
CODIGO = Código da Fiscalização. Requer: pymupdf, pillow.
"""

from __future__ import annotations

import io
import logging
import re
import unicodedata
import zipfile
from pathlib import Path
from typing import Optional

logger = logging.getLogger(__name__)

try:
    import fitz          # PyMuPDF
    FITZ_OK = True
except ImportError:
    FITZ_OK = False

try:
    from PIL import Image as PILImage
    PIL_OK = True
except ImportError:
    PIL_OK = False

ALTURA_CABECALHO_NC   = 120
ALTURA_BUSCA_TEXTO    = 280
ALTURA_TEXTO_ABAIXO   = 350
ALTURA_FAIXA_ESCURA   = 45
Y0_MINIMO_BLOCO       = 66
MARGEM_SUPERIOR       = 4
FOLGA_APOS_FOTO_ANT   = 18

# Fallback se nc_artesp.config não carregar (igual defaults M02)
NC_IMAGE_WIDTH  = 800
NC_IMAGE_HEIGHT = 500
NC_IMAGE_DPI_X  = 222
NC_IMAGE_DPI_Y  = 319


def _cfg_m02_foto_nc() -> tuple[int, int, int, int]:
    try:
        from nc_artesp.config import (
            M02_FOTO_DPI_X,
            M02_FOTO_DPI_Y,
            M02_FOTO_H,
            M02_FOTO_W,
        )
        return (M02_FOTO_W, M02_FOTO_H, M02_FOTO_DPI_X, M02_FOTO_DPI_Y)
    except Exception:
        return (NC_IMAGE_WIDTH, NC_IMAGE_HEIGHT, NC_IMAGE_DPI_X, NC_IMAGE_DPI_Y)


def _cfg_m02_foto_pdf_preview() -> tuple[int, int, int, int]:
    try:
        from nc_artesp.config import (
            M02_FOTO_DPI_X,
            M02_FOTO_DPI_Y,
            M02_FOTO_PDF_H,
            M02_FOTO_PDF_W,
        )
        return (M02_FOTO_PDF_W, M02_FOTO_PDF_H, M02_FOTO_DPI_X, M02_FOTO_DPI_Y)
    except Exception:
        return (480, 202, NC_IMAGE_DPI_X, NC_IMAGE_DPI_Y)


def _resolve_dpi_extracao(dpi: Optional[int]) -> int:
    if dpi is not None:
        return int(dpi)
    try:
        from nc_artesp.config import M02_EXTRACAO_RENDER_DPI
        return int(M02_EXTRACAO_RENDER_DPI)
    except Exception:
        return 150

# Quadros em branco: fração mínima de pixels "não brancos" para considerar como foto real
_UMBRAL_BRANCO = 250
_FRACAO_PIXELS_NAO_BRANCOS_MIN = 0.05  # 5% — quadros com só borda passam a ser filtrados


def _eh_jpg_quase_em_branco(jpg_bytes: bytes) -> bool:
    """True se o conteúdo do JPEG for quase todo branco (página/quadro em branco)."""
    if not PIL_OK or not jpg_bytes or len(jpg_bytes) < 100:
        return False
    try:
        img = PILImage.open(io.BytesIO(jpg_bytes))
        if img.mode in ("RGBA", "P"):
            img = img.convert("RGB")
        w, h = img.size
        if w * h == 0:
            return True
        try:
            resample = PILImage.Resampling.LANCZOS
        except AttributeError:
            resample = getattr(PILImage, "LANCZOS", PILImage.BICUBIC)
        img = img.resize((min(80, w), min(80, h)), resample)
        pixels = list(img.getdata())
        nao_brancos = sum(1 for (r, g, b) in pixels if r < _UMBRAL_BRANCO or g < _UMBRAL_BRANCO or b < _UMBRAL_BRANCO)
        return (nao_brancos / len(pixels)) < _FRACAO_PIXELS_NAO_BRANCOS_MIN
    except Exception:
        return False


def _eh_imagem_embutida_em_branco(doc: "fitz.Document", xref: int) -> bool:
    """True se a imagem embutida for quase toda branca (quadro vazio/placeholder)."""
    if not PIL_OK:
        return False
    try:
        base = doc.extract_image(xref)
        img_bytes = base.get("image")
        if not img_bytes:
            return True
        img = PILImage.open(io.BytesIO(img_bytes))
        if img.mode in ("RGBA", "P"):
            img = img.convert("RGB")
        if img.mode != "RGB":
            return False
        w, h = img.size
        if w * h == 0:
            return True
        try:
            resample = PILImage.Resampling.LANCZOS
        except AttributeError:
            resample = getattr(PILImage, "LANCZOS", PILImage.BICUBIC)
        img = img.resize((min(80, w), min(80, h)), resample)
        pixels = list(img.getdata())
        nao_brancos = sum(1 for (r, g, b) in pixels if r < _UMBRAL_BRANCO or g < _UMBRAL_BRANCO or b < _UMBRAL_BRANCO)
        return (nao_brancos / len(pixels)) < _FRACAO_PIXELS_NAO_BRANCOS_MIN
    except Exception:
        return False


def _check_deps() -> None:
    if not FITZ_OK:
        raise ImportError(
            "PyMuPDF não instalado.\n"
            "Execute: pip install pymupdf"
        )
    if not PIL_OK:
        raise ImportError(
            "Pillow não instalado.\n"
            "Execute: pip install pillow"
        )


def _obter_rects_fotos(page: "fitz.Page") -> list:
    """Retângulos das fotos na página (ordem top→bottom). Exclui quadros em branco (placeholders)."""
    rects = []
    try:
        doc = getattr(page, "parent", None)
        for img in page.get_images():
            xref = img[0]
            if doc and _eh_imagem_embutida_em_branco(doc, xref):
                continue
            for r in page.get_image_rects(xref):
                if r.width > 50 and r.height > 50:
                    rects.append(r)
    except Exception:
        pass
    rects.sort(key=lambda r: (r.y0, r.x0))
    return rects


def _bloco_texto_e_foto(page: "fitz.Page", y0_busca: float,
                         foto_rect: "fitz.Rect",
                         y0_minimo: Optional[float] = None,
                         y1_limite_abaixo: Optional[float] = None) -> "fitz.Rect":
    """Retângulo do bloco: texto acima + foto + texto abaixo. PDF (N).jpg = bloco; nc (N).jpg = foto_rect."""
    if y0_minimo is None:
        y0_minimo = Y0_MINIMO_BLOCO
    x0 = foto_rect.x0
    x1 = foto_rect.x1
    y1 = foto_rect.y1
    y0_final = y0_busca
    try:
        clip = fitz.Rect(0, y0_busca, page.rect.width, y1)
        full = page.get_text("dict", clip=clip)
        blocks = (full.get("blocks", []) if isinstance(full, dict) else []) or []
        for blk in blocks:
            bbox = blk.get("bbox")
            if not bbox:
                continue
            bx0, by0, bx1, by1 = bbox
            if by1 < y0_busca or by0 > y1:
                continue
            if by0 < ALTURA_FAIXA_ESCURA:
                continue
            # Conservação: não incluir cabeçalho da próxima NC (Código Fiscalização: N)
            if by0 >= foto_rect.y0:
                texto_blk = " ".join(
                    s.get("text", "") for line in blk.get("lines", []) for s in line.get("spans", [])
                )
                if re.search(r"C[oó]digo\s+(da\s+)?Fiscaliza[cç][aã]o\s*:\s*\d", texto_blk, re.I):
                    continue
            x0 = min(x0, bx0)
            x1 = max(x1, bx1)
            y0_final = min(y0_final, by0 - MARGEM_SUPERIOR)
        y1_abaixo = min(
            foto_rect.y1 + ALTURA_TEXTO_ABAIXO,
            page.rect.height,
            (y1_limite_abaixo if y1_limite_abaixo is not None else page.rect.height)
        )
        clip_abaixo = fitz.Rect(0, foto_rect.y1, page.rect.width, y1_abaixo)
        full_abaixo = page.get_text("dict", clip=clip_abaixo)
        blocks_abaixo = (full_abaixo.get("blocks", []) if isinstance(full_abaixo, dict) else []) or []
        for blk in blocks_abaixo:
            bbox = blk.get("bbox")
            if not bbox:
                continue
            bx0, by0, bx1, by1 = bbox
            if by0 < foto_rect.y1 - 5:
                continue
            # Não incluir cabeçalho da próxima NC
            texto_blk = " ".join(
                s.get("text", "") for line in blk.get("lines", []) for s in line.get("spans", [])
            )
            if re.search(r"C[oó]digo\s+(da\s+)?Fiscaliza[cç][aã]o\s*:\s*\d", texto_blk, re.I):
                continue
            x0 = min(x0, bx0)
            x1 = max(x1, bx1)
            y1 = max(y1, by1)
    except Exception:
        pass
    y0_final = max(y0_minimo, y0_final, 0)
    return fitz.Rect(x0, y0_final, x1, y1)


def _rect_texto_acima_fotos(page: "fitz.Page", fotos: list) -> Optional["fitz.Rect"]:
    """Faixa horizontal da página do topo até logo acima da primeira foto (altura mín. ~50 pt)."""
    if not fotos:
        return None
    y1 = min(r.y0 for r in fotos) - 6
    y0 = page.rect.y0 + MARGEM_SUPERIOR
    if y1 - y0 < 50:
        return None
    return fitz.Rect(page.rect.x0 + 1, y0, page.rect.x1 - 1, min(y1, page.rect.y1 - 1))


def _renderizar_jpg(page: "fitz.Page", rect: "fitz.Rect", dpi: int = 150) -> bytes:
    """Renderiza retângulo da página como JPEG."""
    clip = rect.intersect(page.rect)
    if clip.is_empty:
        return b""
    pix = page.get_pixmap(dpi=dpi, alpha=False, clip=clip)
    png_bytes = pix.tobytes("png")
    img = PILImage.open(io.BytesIO(png_bytes))
    if img.mode in ("RGBA", "P"):
        img = img.convert("RGB")
    buf = io.BytesIO()
    img.save(buf, "JPEG", quality=92)
    return buf.getvalue()


def _redimensionar_nc_jpg(img_bytes: bytes) -> bytes:
    """nc (COD).jpg: M02_FOTO_W×H e M02_FOTO_DPI (ARTESP = Artemig)."""
    w, h, dx, dy = _cfg_m02_foto_nc()
    img = PILImage.open(io.BytesIO(img_bytes))
    if img.mode in ("RGBA", "P"):
        img = img.convert("RGB")
    try:
        resample = PILImage.Resampling.LANCZOS
    except AttributeError:
        resample = getattr(PILImage, "LANCZOS", PILImage.BICUBIC)
    img = img.resize((w, h), resample)
    buf = io.BytesIO()
    img.save(buf, "JPEG", quality=92, dpi=(dx, dy))
    return buf.getvalue()


def _redimensionar_pdf_ou_texto_jpg(img_bytes: bytes) -> bytes:
    """PDF (COD).jpg / Texto (COD).jpg: M02_FOTO_PDF_W×H e mesmo DPI M02."""
    pw, ph, dx, dy = _cfg_m02_foto_pdf_preview()
    img = PILImage.open(io.BytesIO(img_bytes))
    if img.mode in ("RGBA", "P"):
        img = img.convert("RGB")
    try:
        resample = PILImage.Resampling.LANCZOS
    except AttributeError:
        resample = getattr(PILImage, "LANCZOS", PILImage.BICUBIC)
    img = img.resize((pw, ph), resample)
    buf = io.BytesIO()
    img.save(buf, "JPEG", quality=92, dpi=(dx, dy))
    return buf.getvalue()


def _eh_codigo_fiscalizacao_valido(val: str) -> bool:
    """Aceita valor numérico (código fiscalização). Rejeita Lote, Grau, etc."""
    if not val or not isinstance(val, str):
        return False
    s = val.strip()
    if not s or s.upper().startswith("LOTE"):
        return False
    if s.isdigit():
        return True
    if re.match(r"^\d+[\-]?\d*$", s):
        return True
    return False


def _extrair_codigo_por_blocos(page: "fitz.Page", clip_rect: "fitz.Rect") -> str:
    """Fallback PDF em tabela: Código da Fiscalização e número na mesma linha (mesmo y)."""
    try:
        full = page.get_text("dict", clip=clip_rect)
        blocks = (full.get("blocks", []) if isinstance(full, dict) else []) or []
        y_rotulo = None
        candidatos = []
        for blk in blocks:
            for line in blk.get("lines", []):
                for span in line.get("spans", []):
                    t = (span.get("text") or "").strip()
                    bbox = span.get("bbox")
                    if not bbox or len(bbox) < 4:
                        continue
                    y = bbox[1]
                    if re.search(r"C[oó]digo\s+da\s+Fiscaliza[cç][aã]o\s*:?", t, re.I):
                        y_rotulo = y
                    elif re.search(r"^C[oó]digo\s*:?\s*$", t, re.I) or (t.lower().strip() in ("codigo", "código", "codigo:", "código:") and "fiscaliza" not in t.lower()):
                        y_rotulo = y
                    if t.isdigit() and len(t) >= 4:
                        candidatos.append((y, t))
        if y_rotulo is None or not candidatos:
            return ""
        for y, num in candidatos:
            if abs(y - y_rotulo) < 20 and _eh_codigo_fiscalizacao_valido(num):
                return num
        if candidatos and y_rotulo is not None:
            melhor = min(candidatos, key=lambda c: abs(c[0] - y_rotulo))
            if _eh_codigo_fiscalizacao_valido(melhor[1]):
                return melhor[1]
        return ""
    except Exception:
        return ""


def _texto_pagina_em_ordem_leitura(page: "fitz.Page", clip_rect: "fitz.Rect") -> str:
    """Texto da página em ordem de leitura (blocos por y, x), como no analisar_pdf."""
    try:
        blocos = page.get_text("blocks", clip=clip_rect)
        if not blocos:
            return page.get_text("text", clip=clip_rect) or ""
        blocos.sort(key=lambda b: (round(b[1], 0), round(b[0], 0)))
        return "\n".join((b[4] or "").strip() for b in blocos if (b[4] or "").strip())
    except Exception:
        return page.get_text("text", clip=clip_rect) or ""


def _extrair_codigo_nc(page: "fitz.Page", bloco_rect: "fitz.Rect") -> str:
    """Extrai código da fiscalização para nomear fotos. Conservação: Lote: 896643 → 896643."""
    def _rejeitar_lote(texto: str, val: str) -> str:
        if not val:
            return val
        v = val.strip().upper()
        if v != "LOTE" and not v.startswith("LOTE"):
            return val
        m = re.search(r"Lote\s*:\s*(\S+)", texto, re.IGNORECASE)
        return (m.group(1) or "").strip() if m else ""

    def _nunca_lote(s: str) -> str:
        if not s or s.strip().upper().startswith("LOTE"):
            return ""
        return s.strip()

    try:
        texto = page.get_text("text", clip=bloco_rect)
        # Para página inteira, usar ordem de leitura (blocos) como no analisar_pdf
        if bloco_rect.get_area() >= 0.85 * (page.rect.get_area() or 1):
            texto_ordenado = _texto_pagina_em_ordem_leitura(page, bloco_rect)
            if texto_ordenado.strip():
                texto = texto_ordenado
        if len((texto or "").strip()) < 30:
            try:
                from nc_artesp.pdf_ocr import texto_de_pagina_ocr
                ocr = texto_de_pagina_ocr(page, rect=bloco_rect, dpi=200)
                if ocr:
                    texto = ocr
            except Exception:
                pass
        m = re.search(
            r'C[oó]digo\s+da\s+Fiscaliza[cç][aã]o\s*:\s*(\S+)',
            texto, re.IGNORECASE
        )
        if m:
            val = _nunca_lote(_rejeitar_lote(texto, m.group(1).strip()))
            if val and _eh_codigo_fiscalizacao_valido(val):
                return val
        m = re.search(
            r'C[oó]digo\s+Fiscaliza[cç][aã]o\s*:\s*(\S+)',
            texto, re.IGNORECASE
        )
        if m:
            val = _nunca_lote(_rejeitar_lote(texto, m.group(1).strip()))
            if val and _eh_codigo_fiscalizacao_valido(val):
                return val
        m = re.search(
            r'C[oó]digo\s+Fiscaliza[cç][aã]o:\s*Lote:\s*(\S+)',
            texto, re.IGNORECASE
        )
        if m:
            val = _nunca_lote(m.group(1).strip())
            if val and _eh_codigo_fiscalizacao_valido(val):
                return val
        # Legenda só "Código" (sem "da Fiscalização") — mesmo padrão numérico
        m = re.search(
            r'C[oó]digo(?!\s+da)(?!\s+Fiscaliza)\s*:\s*(\S+)',
            texto, re.IGNORECASE
        )
        if m:
            val = _nunca_lote(_rejeitar_lote(texto, m.group(1).strip()))
            if val and _eh_codigo_fiscalizacao_valido(val):
                return val
        m = re.search(
            r'C[oó]digo(?!\s+da)(?!\s+Fiscaliza)\s+(\d{4,})',
            texto, re.IGNORECASE
        )
        if m:
            val = m.group(1).strip()
            if _eh_codigo_fiscalizacao_valido(val):
                return val
        # Número antes da legenda (ex.: "896643 Código")
        m = re.search(
            r'(\d{4,})\s+C[oó]digo(?!\s+da)(?!\s+Fiscaliza)',
            texto, re.IGNORECASE
        )
        if m and _eh_codigo_fiscalizacao_valido(m.group(1).strip()):
            return m.group(1).strip()
        # Layout emergencial: mesma lógica do analisar_pdf — acha linha só "Código", depois primeira linha 5+ dígitos
        if re.search(r"\bC[oó]digo\b", texto, re.IGNORECASE) and re.search(r"\d{5,}", texto):
            linhas = [ln.strip() for ln in texto.splitlines()]
            idx_codigo = next(
                (i for i, ln in enumerate(linhas) if re.match(r"^\s*C[oó]digo\s*$", ln, re.IGNORECASE)),
                -1,
            )
            if idx_codigo >= 0:
                for ln in linhas[idx_codigo + 1 : min(idx_codigo + 15, len(linhas))]:
                    if re.match(r"^Lote\s*:?\s*", ln, re.IGNORECASE) or re.match(r"^Data\s+da\s+Constata", ln, re.IGNORECASE):
                        break
                    if re.match(r"^\s*\d{5,}\s*$", ln) and _eh_codigo_fiscalizacao_valido(ln.strip()):
                        return ln.strip()
            # Fallback: primeira linha só 5+ dígitos antes de Lote/Data
            for ln in linhas:
                if re.match(r"^Lote\s*:?", ln, re.IGNORECASE) or re.match(r"^Data\s+da\s+Constata", ln, re.IGNORECASE):
                    break
                if re.match(r"^\s*\d{5,}\s*$", ln) and _eh_codigo_fiscalizacao_valido(ln.strip()):
                    return ln.strip()
        # Sem linha "Código" explícita: primeira linha com 5+ dígitos (layout alternativo)
        if re.search(r"\d{5,}", texto):
            for ln in [ln.strip() for ln in texto.splitlines()]:
                if re.match(r"^Lote\s*:?", ln, re.IGNORECASE) or re.match(r"^Data\s+da\s+Constata", ln, re.IGNORECASE):
                    break
                if re.match(r"^\s*\d{5,}\s*$", ln) and _eh_codigo_fiscalizacao_valido(ln):
                    return ln
        if "Fiscaliza" in texto or "fiscaliza" in texto or "codigo" in texto.lower():
            cod = _extrair_codigo_por_blocos(page, bloco_rect)
            if cod:
                return cod
            cod = _extrair_codigo_por_blocos(page, page.rect)
            if cod:
                return cod
    except Exception:
        pass
    return ""


def _codigo_estilo_ma(codigo: str) -> bool:
    """Código no padrão MA (ponto e letras)."""
    if not codigo or not isinstance(codigo, str):
        return False
    s = str(codigo).strip()
    return "." in s and any(c.isalpha() for c in s)


def _nome_arquivo_safe(s: str) -> str:
    """Garante string encodável em Latin-1 para nomes no ZIP/Content-Disposition."""
    if not s:
        return s
    nfd = unicodedata.normalize("NFD", s)
    sem_comb = "".join(c for c in nfd if unicodedata.category(c) != "Mn")
    try:
        return sem_comb.encode("latin-1").decode("latin-1")
    except UnicodeEncodeError:
        return sem_comb.encode("latin-1", "replace").decode("latin-1")


def _formatar_codigo_arquivo(codigo: str, num_digitos: int = 5) -> str:
    """Código para nome do arquivo (zeros à esquerda); sanitizado para Latin-1."""
    s = (codigo or "").strip()
    try:
        n = int(s)
        return str(n).zfill(num_digitos)
    except (ValueError, TypeError):
        return _nome_arquivo_safe(s)


def extrair_imagens_pdf(pdf_path: str,
                         pasta_saida: Optional[str] = None,
                         pasta_saida_nc: Optional[str] = None,
                         pasta_saida_pdf: Optional[str] = None,
                         dpi: Optional[int] = None,
                         nc_global_start: int = 0,
                         nomear_por_indice_fiscalizacao: bool = False,
                         pasta_unica: bool = False) -> list:
    """Extrai Texto (COD).jpg, PDF (COD).jpg e nc (COD).jpg. Se pasta_unica, tudo na mesma pasta."""
    _check_deps()
    dpi = _resolve_dpi_extracao(dpi)
    pdf_path = Path(pdf_path).resolve()
    if not pdf_path.exists():
        raise FileNotFoundError(f"PDF não encontrado: {pdf_path}")

    if pasta_unica:
        base = Path(pasta_saida).resolve() if pasta_saida else pdf_path.parent
        base.mkdir(parents=True, exist_ok=True)
        p_nc = p_pdf = base
    else:
        usar_duas = bool(pasta_saida_nc and pasta_saida_pdf)
        if usar_duas:
            p1 = Path(pasta_saida_nc).resolve()
            p2 = Path(pasta_saida_pdf).resolve()
            if p1 == p2:
                p_nc = p1 / "nc"
                p_pdf = p1 / "PDF"
                p_nc.mkdir(parents=True, exist_ok=True)
                p_pdf.mkdir(parents=True, exist_ok=True)
            else:
                p_nc, p_pdf = p1, p2
                p_nc.mkdir(parents=True, exist_ok=True)
                p_pdf.mkdir(parents=True, exist_ok=True)
        else:
            base = Path(pasta_saida).resolve() if pasta_saida else pdf_path.parent
            base.mkdir(parents=True, exist_ok=True)
            p_nc = base / "nc"
            p_pdf = base / "PDF"
            p_nc.mkdir(parents=True, exist_ok=True)
            p_pdf.mkdir(parents=True, exist_ok=True)

    salvos    = []
    nc_global = nc_global_start
    nomes_usados: set[str] = set()
    ultimo_codigo: Optional[str] = None
    usar_indice = nomear_por_indice_fiscalizacao
    codigos_doc: list[str] = []
    try:
        from nc_artesp.modulos.analisar_pdf_nc import parse_pdf_nc
        ncs = parse_pdf_nc(pdf_path.read_bytes())
        codigos_doc = [(nc.codigo or "").strip() for nc in ncs if (nc.codigo or "").strip()]
    except Exception as e:
        logger.debug("Fallback códigos do documento: %s", e)
    codigo_idx = 0
    doc = fitz.open(str(pdf_path))

    def _nome_unico(base_nome: str) -> str:
        nome = base_nome
        n = 1
        while nome in nomes_usados:
            stem = base_nome.rsplit(".", 1)[0]
            nome = f"{stem}_{n}.jpg"
            n += 1
        nomes_usados.add(nome)
        return nome

    try:
        for page_num in range(len(doc)):
            page = doc[page_num]
            if usar_indice:
                ultimo_codigo = None
            # Uma imagem PDF (texto+foto) por página; página em branco já não gera bloco
            pdf_imagem_ja_escrita_esta_pagina = False
            try:
                r_fotos = _obter_rects_fotos(page)
            except Exception as e:
                logger.debug("Página %s: obter rects fotos: %s", page_num + 1, e)
                r_fotos = []

            texto_pagina = ""
            try:
                texto_pagina = page.get_text("text", clip=page.rect) or ""
            except Exception:
                pass
            eh_ma = "Código da Fiscalização" in texto_pagina or "Meio Ambiente" in texto_pagina or "codigo da fiscalização" in (texto_pagina or "").lower()

            # Cabeçalho: topo da página; se tiver "Código" ou "Código da Fiscalização" = início de NC; senão, página só com fotos = continuação da NC anterior
            texto_cabecalho = ""
            try:
                rect_cabecalho = fitz.Rect(page.rect.x0, page.rect.y0, page.rect.x1, page.rect.y0 + ALTURA_CABECALHO_NC)
                texto_cabecalho = (page.get_text("text", clip=rect_cabecalho) or "").strip()
            except Exception:
                pass
            trecho_ini = ((texto_cabecalho or "") + "\n" + (texto_pagina or "")[:1800])[:2500]
            tem_cabecalho_nc = bool(
                re.search(r"C[oó]digo\s+(da\s+)?Fiscaliza[cç][aã]o", texto_cabecalho, re.I)
                or re.search(r"^\s*C[oó]digo\s*$", texto_cabecalho, re.M | re.I)
                or re.search(
                    r"NOTIFICA[ÇC][AÃ]O|N[ºo°]?\s*da\s*CONSOL|Tipo\s*:\s*QID|Constata[cç][aã]o",
                    trecho_ini,
                    re.I,
                )
            )
            pagina_continuacao = not tem_cabecalho_nc

            if not r_fotos:
                # Página sem imagens embutidas: só tratar como "uma foto" se o conteúdo não for em branco
                jpg_teste = _renderizar_jpg(page, page.rect, dpi)
                if jpg_teste and not _eh_jpg_quase_em_branco(jpg_teste):
                    r_fotos = [page.rect]
                    blocos = [(page.rect, page.rect)]
                else:
                    r_fotos = []
                    blocos = []
            else:
                blocos = []
                for i, r in enumerate(r_fotos):
                    try:
                        y0_busca = max(0, r.y0 - ALTURA_BUSCA_TEXTO)
                        if i > 0:
                            y0_busca = max(y0_busca, r_fotos[i - 1].y1 + FOLGA_APOS_FOTO_ANT)
                        y0_min = Y0_MINIMO_BLOCO if i == 0 else y0_busca
                        if eh_ma:
                            y1_limite = r_fotos[i + 1].y0 - 1 if i + 1 < len(r_fotos) else None
                        else:
                            y1_limite = r.y1
                        bloco = _bloco_texto_e_foto(page, y0_busca, r, y0_minimo=y0_min, y1_limite_abaixo=y1_limite)
                        blocos.append((bloco, r))
                    except Exception as e:
                        logger.debug("Página %s bloco %s: %s", page_num + 1, i, e)
                        blocos.append((r, r))

            def flush_grupo(bloco_uniao: "fitz.Rect", fotos: list, cod: str):
                nonlocal pdf_imagem_ja_escrita_esta_pagina
                if bloco_uniao is None or not cod:
                    return
                try:
                    # Uma imagem PDF (texto+foto) por página; nc = todas as fotos. Página em branco não gera bloco.
                    fotos_unicas = []
                    vistos_rect = set()
                    for fr in fotos:
                        key = (round(fr.x0, 2), round(fr.y0, 2), round(fr.x1, 2), round(fr.y1, 2))
                        if key not in vistos_rect:
                            vistos_rect.add(key)
                            fotos_unicas.append(fr)
                    so_foto = (
                        len(fotos_unicas) == 1
                        and abs(bloco_uniao.x0 - fotos_unicas[0].x0) < 1
                        and abs(bloco_uniao.y0 - fotos_unicas[0].y0) < 1
                        and abs(bloco_uniao.x1 - fotos_unicas[0].x1) < 1
                        and abs(bloco_uniao.y1 - fotos_unicas[0].y1) < 1
                    )
                    clip_pdf = None
                    if not pagina_continuacao and fotos_unicas:
                        if so_foto:
                            fr0 = fotos_unicas[0]
                            clip_pdf = fitz.Rect(
                                page.rect.x0 + 1,
                                page.rect.y0 + MARGEM_SUPERIOR,
                                page.rect.x1 - 1,
                                min(page.rect.y1 - 1, fr0.y1 + 25),
                            )
                        elif bloco_uniao is not None:
                            clip_pdf = bloco_uniao
                    clip_txt = None
                    if not pagina_continuacao and fotos_unicas:
                        clip_txt = _rect_texto_acima_fotos(page, fotos_unicas)
                    if (
                        pasta_unica
                        and
                        clip_txt
                        and not pdf_imagem_ja_escrita_esta_pagina
                        and clip_txt.get_area() > 400
                    ):
                        jpg_txt = _renderizar_jpg(page, clip_txt, dpi)
                        if jpg_txt and not _eh_jpg_quase_em_branco(jpg_txt):
                            nome_t = _nome_unico(f"Texto ({cod}).jpg")
                            (p_pdf / nome_t).write_bytes(_redimensionar_pdf_ou_texto_jpg(jpg_txt))
                            salvos.append(str(p_pdf / nome_t))
                    if (
                        clip_pdf
                        and not pdf_imagem_ja_escrita_esta_pagina
                        and clip_pdf.get_area() > 500
                    ):
                        jpg_pdf = _renderizar_jpg(page, clip_pdf, dpi)
                        if jpg_pdf and not _eh_jpg_quase_em_branco(jpg_pdf):
                            nome = _nome_unico(f"PDF ({cod}).jpg")
                            (p_pdf / nome).write_bytes(_redimensionar_pdf_ou_texto_jpg(jpg_pdf))
                            salvos.append(str(p_pdf / nome))
                            pdf_imagem_ja_escrita_esta_pagina = True
                    for fr in fotos_unicas:
                        jpg_foto = _renderizar_jpg(page, fr, dpi)
                        if jpg_foto and not _eh_jpg_quase_em_branco(jpg_foto):
                            jpg_foto = _redimensionar_nc_jpg(jpg_foto)
                            nome = _nome_unico(f"nc ({cod}).jpg")
                            (p_nc / nome).write_bytes(jpg_foto)
                            salvos.append(str(p_nc / nome))
                except Exception as e:
                    logger.warning("Página %s flush_grupo (%s): %s", page_num + 1, cod, e)

            grupo_rect = None
            grupo_fotos = []
            grupo_codigo = None

            try:
                for bloco_rect, foto_rect in blocos:
                    codigo_extraido = _extrair_codigo_nc(page, bloco_rect)
                    if not codigo_extraido:
                        codigo_extraido = _extrair_codigo_nc(page, page.rect)
                    if usar_indice:
                        if codigo_extraido or ultimo_codigo is None:
                            nc_global += 1
                            ultimo_codigo = str(nc_global).zfill(5)
                        codigo_nome = ultimo_codigo
                    else:
                        nc_global += 1
                        if codigo_extraido:
                            codigo = codigo_extraido
                            ultimo_codigo = codigo
                            if codigo_idx < len(codigos_doc):
                                codigo_idx += 1
                        elif codigo_idx < len(codigos_doc) and _eh_codigo_fiscalizacao_valido(codigos_doc[codigo_idx]):
                            codigo = codigos_doc[codigo_idx]
                            codigo_idx += 1
                            ultimo_codigo = codigo
                        else:
                            codigo = ultimo_codigo if ultimo_codigo else str(nc_global)
                        codigo_nome = _formatar_codigo_arquivo(codigo)
                    if not codigo_nome or codigo_nome.upper().startswith("LOTE"):
                        codigo_nome = str(nc_global)

                    if grupo_codigo is not None and grupo_codigo != codigo_nome:
                        flush_grupo(grupo_rect, grupo_fotos, grupo_codigo)
                        grupo_rect = None
                        grupo_fotos = []
                    grupo_codigo = codigo_nome
                    if grupo_rect is None:
                        grupo_rect = bloco_rect
                    else:
                        grupo_rect = fitz.Rect(
                            min(grupo_rect.x0, bloco_rect.x0),
                            min(grupo_rect.y0, bloco_rect.y0),
                            max(grupo_rect.x1, bloco_rect.x1),
                            max(grupo_rect.y1, bloco_rect.y1),
                        )
                    grupo_fotos.append(foto_rect)

                if grupo_codigo is not None:
                    flush_grupo(grupo_rect, grupo_fotos, grupo_codigo)
            except Exception as e:
                logger.warning("Página %s estrutura diferente, usando página inteira: %s", page_num + 1, e)
                jpg_teste = _renderizar_jpg(page, page.rect, dpi)
                if jpg_teste and not _eh_jpg_quase_em_branco(jpg_teste):
                    nc_global += 1
                    cod_fallback = str(nc_global).zfill(5)
                    flush_grupo(page.rect, [page.rect], cod_fallback)
                    ultimo_codigo = cod_fallback
    finally:
        doc.close()

    return salvos


def extrair_pdf_para_zip(pdf_bytes: bytes, dpi: Optional[int] = None,
                         nomear_por_indice_fiscalizacao: bool = False,
                         pasta_unica: bool = False) -> tuple[bytes, int]:
    """PDF → ZIP. pasta_unica: Texto/PDF/nc na mesma pasta (ex.: Artemig lote 50)."""
    _check_deps()
    import tempfile
    with tempfile.TemporaryDirectory() as tmpdir:
        pdf_path = Path(tmpdir) / "upload.pdf"
        pdf_path.write_bytes(pdf_bytes)
        pasta_saida = Path(tmpdir) / "saida"
        pasta_saida.mkdir()
        salvos = extrair_imagens_pdf(
            str(pdf_path),
            pasta_saida=str(pasta_saida),
            dpi=dpi,
            nomear_por_indice_fiscalizacao=nomear_por_indice_fiscalizacao,
            pasta_unica=pasta_unica,
        )
        pasta_raiz = pasta_saida.resolve()
        buf = io.BytesIO()
        with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
            for f in salvos:
                fp = Path(f)
                try:
                    arc = str(fp.relative_to(pasta_raiz)).replace("\\", "/")
                except ValueError:
                    arc = fp.name
                zf.write(fp, arc)
        n_ncs = len([f for f in salvos if Path(f).name.startswith("PDF (")])
        return buf.getvalue(), n_ncs
