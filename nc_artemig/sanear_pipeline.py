"""Saneamento único para NCs Artemig lote 50 (QID, CONSOL) antes de relatórios e export Kcor."""

from __future__ import annotations

import re
from typing import Any


def norm_lote_numero(lote: Any) -> str:
    """Primeiro grupo de dígitos normalizado (ex.: '050' → '50'). Vazio se não houver."""
    s = str(lote if lote is not None else "").strip()
    m = re.search(r"\d+", s)
    if not m:
        return ""
    try:
        return str(int(m.group(0)))
    except ValueError:
        return m.group(0)


def relatorio_deve_tratar_artemig(lote_selecionado: Any, ncs: list[Any]) -> bool:
    """True se o relatório XLSX deve aplicar regras Artemig lote 50 antes do preenchimento."""
    if norm_lote_numero(lote_selecionado) == "50":
        return True
    return any(norm_lote_numero(getattr(n, "lote", None) or "") == "50" for n in ncs)


def _identificador_sem_whitespace(val: Any) -> str:
    try:
        from nc_artemig.texto_pdf import identificador_pdf_sem_whitespace

        return identificador_pdf_sem_whitespace(val)
    except ImportError:
        return re.sub(r"\s+", "", str(val if val is not None else "").strip())


def sanear_ncs_lote50_consol(ncs: list[Any], *, forcar_todas: bool = False) -> None:
    """
    Força lote 50 e tipo QID nas NCs alvo; reduz ruído de espaços em campos CONSOL/SH.
    `forcar_todas=True` aplica a todas as NCs da lista (análise/export já filtrados por contexto 50).
    """
    if not ncs:
        return

    def _incluir(nc: Any) -> bool:
        if forcar_todas:
            return True
        return norm_lote_numero(getattr(nc, "lote", None) or "") == "50"

    for nc in ncs:
        if not _incluir(nc):
            continue
        nc.lote = "50"
        nc.tipo_artemig = "QID"
        if hasattr(nc, "num_consol"):
            raw = getattr(nc, "num_consol", None)
            if raw is not None and str(raw).strip():
                nc.num_consol = _identificador_sem_whitespace(raw)
        if hasattr(nc, "sh_artemig"):
            sh = getattr(nc, "sh_artemig", None)
            if isinstance(sh, str) and sh.strip():
                nc.sh_artemig = re.sub(r"\s+", " ", sh.strip())
