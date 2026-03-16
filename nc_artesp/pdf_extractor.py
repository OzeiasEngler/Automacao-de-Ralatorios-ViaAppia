"""
Extração de imagens do PDF de NC Constatação Artesp: PDF (CODIGO).jpg (texto+foto) e nc (CODIGO).jpg (só foto).
CODIGO = Código da Fiscalização (nunca Lote). Requer: pymupdf, pillow.
"""

from __future__ import annotations

import io
import re
import zipfile
from pathlib import Path
from typing import Optional

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

# nc (N).jpg na extração: 800×500 px, 222×319 DPI
NC_IMAGE_WIDTH  = 800
NC_IMAGE_HEIGHT = 500
NC_IMAGE_DPI_X  = 222
NC_IMAGE_DPI_Y  = 319


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
    """Retângulos das fotos na página (ordem top→bottom)."""
    rects = []
    try:
        for img in page.get_images():
            xref = img[0]
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
        for blk in page.get_text("dict", clip=clip).get("blocks", []):
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
        for blk in page.get_text("dict", clip=clip_abaixo).get("blocks", []):
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
    """Saída nc (N).jpg: 800×500 px, DPI 222×319."""
    img = PILImage.open(io.BytesIO(img_bytes))
    if img.mode in ("RGBA", "P"):
        img = img.convert("RGB")
    img = img.resize((NC_IMAGE_WIDTH, NC_IMAGE_HEIGHT), PILImage.LANCZOS)
    buf = io.BytesIO()
    img.save(buf, "JPEG", quality=92, dpi=(NC_IMAGE_DPI_X, NC_IMAGE_DPI_Y))
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
        y_rotulo = None
        candidatos = []
        for blk in full.get("blocks", []):
            for line in blk.get("lines", []):
                for span in line.get("spans", []):
                    t = (span.get("text") or "").strip()
                    bbox = span.get("bbox")
                    if not bbox or len(bbox) < 4:
                        continue
                    y = bbox[1]
                    if re.search(r"C[oó]digo\s+da\s+Fiscaliza[cç][aã]o\s*:?", t, re.I):
                        y_rotulo = y
                    if t.isdigit() and len(t) >= 5:
                        candidatos.append((y, t))
        if y_rotulo is None or not candidatos:
            return ""
        for y, num in candidatos:
            if abs(y - y_rotulo) < 15 and _eh_codigo_fiscalizacao_valido(num):
                return num
        return ""
    except Exception:
        return ""


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
        if "Fiscaliza" in texto or "fiscaliza" in texto:
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


def _formatar_codigo_arquivo(codigo: str, num_digitos: int = 5) -> str:
    """Código para nome do arquivo (zeros à esquerda)."""
    s = (codigo or "").strip()
    try:
        n = int(s)
        return str(n).zfill(num_digitos)
    except (ValueError, TypeError):
        return s


def extrair_imagens_pdf(pdf_path: str,
                         pasta_saida: Optional[str] = None,
                         pasta_saida_nc: Optional[str] = None,
                         pasta_saida_pdf: Optional[str] = None,
                         dpi: int = 150,
                         nc_global_start: int = 0,
                         nomear_por_indice_fiscalizacao: bool = False) -> list:
    """Extrai PDF (CODIGO).jpg (texto+foto) e nc (CODIGO).jpg (só foto). Retorna lista de caminhos."""
    _check_deps()
    pdf_path = Path(pdf_path).resolve()
    if not pdf_path.exists():
        raise FileNotFoundError(f"PDF não encontrado: {pdf_path}")

    usar_duas = pasta_saida_nc and pasta_saida_pdf
    if usar_duas:
        p_nc  = Path(pasta_saida_nc).resolve()
        p_pdf = Path(pasta_saida_pdf).resolve()
        p_nc.mkdir(parents=True, exist_ok=True)
        p_pdf.mkdir(parents=True, exist_ok=True)
    else:
        base = Path(pasta_saida).resolve() if pasta_saida else pdf_path.parent
        base.mkdir(parents=True, exist_ok=True)
        p_nc = p_pdf = base

    salvos    = []
    nc_global = nc_global_start
    nomes_usados: set[str] = set()
    ultimo_codigo: Optional[str] = None
    usar_indice = nomear_por_indice_fiscalizacao
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
            ultimo_codigo = None
            r_fotos = _obter_rects_fotos(page)

            texto_pagina = ""
            try:
                texto_pagina = page.get_text("text", clip=page.rect)
            except Exception:
                pass
            eh_ma = "Código da Fiscalização" in texto_pagina or "Meio Ambiente" in texto_pagina or "codigo da fiscalização" in texto_pagina.lower()

            if not r_fotos:
                r_fotos = [page.rect]
                blocos = [(page.rect, page.rect)]
            else:
                blocos = []
                for i, r in enumerate(r_fotos):
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

            def flush_grupo(bloco_uniao: "fitz.Rect", fotos: list, cod: str):
                if bloco_uniao is None or not cod:
                    return
                jpg_pdf = _renderizar_jpg(page, bloco_uniao, dpi)
                if jpg_pdf:
                    nome = _nome_unico(f"PDF ({cod}).jpg")
                    (p_nc / nome).write_bytes(jpg_pdf)
                    salvos.append(str(p_nc / nome))
                for fr in fotos:
                    jpg_foto = _renderizar_jpg(page, fr, dpi)
                    if jpg_foto:
                        jpg_foto = _redimensionar_nc_jpg(jpg_foto)
                        nome = _nome_unico(f"nc ({cod}).jpg")
                        (p_pdf / nome).write_bytes(jpg_foto)
                        salvos.append(str(p_pdf / nome))

            grupo_rect = None
            grupo_fotos = []
            grupo_codigo = None

            for bloco_rect, foto_rect in blocos:
                codigo_extraido = _extrair_codigo_nc(page, bloco_rect)
                if not codigo_extraido and eh_ma:
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
    finally:
        doc.close()

    return salvos


def extrair_pdf_para_zip(pdf_bytes: bytes, dpi: int = 150,
                         nomear_por_indice_fiscalizacao: bool = False) -> tuple[bytes, int]:
    """PDF em bytes → ZIP com PDF (CODIGO).jpg e nc (CODIGO).jpg. Retorna (zip_bytes, n_ncs)."""
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
        )
        buf = io.BytesIO()
        with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
            for f in salvos:
                zf.write(f, Path(f).name)
        n_ncs = len([f for f in salvos if Path(f).name.startswith("PDF (")])
        return buf.getvalue(), n_ncs
