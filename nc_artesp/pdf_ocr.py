"""
OCR opcional para páginas de PDF sem camada de texto (escaneados).
Quando get_text() retorna vazio, renderiza a página como imagem e usa Tesseract.
Requer: pip install pytesseract e Tesseract instalado no sistema (https://github.com/tesseract-ocr/tesseract).
"""

from __future__ import annotations

import io
import logging
from typing import TYPE_CHECKING, Optional

logger = logging.getLogger(__name__)

if TYPE_CHECKING:
    import fitz

try:
    import pytesseract
    PYTESSERACT_OK = True
except ImportError:
    PYTESSERACT_OK = False

try:
    from PIL import Image as PILImage
    PIL_OK = True
except ImportError:
    PIL_OK = False


def texto_de_pagina_ocr(
    page: "fitz.Page",
    rect: Optional["fitz.Rect"] = None,
    dpi: int = 200,
    lang: str = "por",
) -> str:
    """
    Extrai texto da página via OCR (Tesseract). Use quando get_text() retornar vazio.
    Retorna "" se pytesseract/PIL não estiverem disponíveis ou Tesseract falhar.
    """
    if not PYTESSERACT_OK or not PIL_OK:
        return ""
    try:
        clip = rect if rect is not None else page.rect
        pix = page.get_pixmap(dpi=dpi, alpha=False, clip=clip)
        img = PILImage.open(io.BytesIO(pix.tobytes("png")))
        if img.mode != "RGB":
            img = img.convert("RGB")
        text = pytesseract.image_to_string(img, lang=lang or "por").strip()
        return text or ""
    except Exception as e:
        logger.debug("OCR página: %s", e)
        return ""
