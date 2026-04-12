"""Normalização de texto extraído de PDF (NBSP, NFKC, espaços) para Artemig / Nas01 / Kcor."""

from __future__ import annotations

import re
import unicodedata
from typing import Any

_NBSP = "\u00a0"


def identificador_pdf_sem_whitespace(val: Any) -> str:
    """Códigos e Nº Consol: remove todo whitespace (PDF costuma inserir NBSP e quebras)."""
    if val is None:
        return ""
    s = unicodedata.normalize("NFKC", str(val)).strip()
    s = s.replace(_NBSP, "")
    return re.sub(r"\s+", "", s)


def limpeza_profunda(s: str) -> str:
    """Uma linha lógica: NFKC, NBSP → espaço, ``\\s+`` → um espaço."""
    if not s:
        return ""
    t = unicodedata.normalize("NFKC", str(s)).replace(_NBSP, " ")
    t = t.replace("\r", " ").replace("\n", " ")
    return re.sub(r"\s+", " ", t).strip()


def colapsar_espacos_pdf(s: str, multiline: bool = False) -> str:
    """Colapsa espaços parasitas do PDF; opcionalmente preserva quebras entre linhas."""
    if s is None:
        return ""
    raw = unicodedata.normalize("NFKC", str(s)).replace(_NBSP, " ")
    if not multiline:
        raw = raw.replace("\r", " ").replace("\n", " ")
        return re.sub(r"\s+", " ", raw).strip()
    lines = [re.sub(r"\s+", " ", ln).strip() for ln in raw.split("\n")]
    return "\n".join(lines)


def limpeza_linha_excel_pdf(ln: str) -> str:
    """Segmento colado em célula (col. T): sem CR/LF/tab no meio da «linha»."""
    t = limpeza_profunda(ln or "")
    return re.sub(r"[\r\n\t]+", "", t)


def limpeza_multilinha_excel_pdf(texto: str) -> str:
    """Observação multilinha antes do corte final para célula Excel."""
    if not texto:
        return ""
    t = unicodedata.normalize("NFKC", str(texto)).replace(_NBSP, " ")
    t = t.replace("\r\n", "\n").replace("\r", "\n")
    lines = [re.sub(r"\s+", " ", ln).strip() for ln in t.split("\n")]
    return "\n".join(x for x in lines if x).strip()
