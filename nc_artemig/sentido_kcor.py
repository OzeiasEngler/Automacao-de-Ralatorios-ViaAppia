"""
Regra de sentido da macro Nas01_Gerar_Plan_Exportar_Kcor (planilha → Kcor).
CRESCENTE/DECRESCENTE/AMBOS convertidos conforme a rodovia (MG-050 vs BR-265/491).
"""

from __future__ import annotations

import re


def _familia_rodovia_artemig(rodovia: str) -> str | None:
    """'MG050' | 'BR265' | 'BR491' ou None se não for malha Artemig."""
    x = re.sub(r"[\s._]", "", (rodovia or "").upper()).replace("-", "")
    if "MG050" in x:
        return "MG050"
    if "BR265" in x:
        return "BR265"
    if "BR491" in x:
        return "BR491"
    return None


def sentido_artemig_para_kcor(rodovia: str, sentido_pdf: str) -> str:
    """
    MG-050: AMBOS → Oeste - SP/Leste - BH; CRESCENTE → Oeste - SP; DECRESCENTE → Leste - BH.
    BR-265 / BR-491: AMBOS → Sul - SP/Norte - BH; CRESCENTE → Norte - BH; DECRESCENTE → Sul - SP.
    Outros textos (já normalizados, vazios) → devolvidos sem alteração.
    """
    fam = _familia_rodovia_artemig(rodovia)
    if not fam:
        return (sentido_pdf or "").strip()
    s = (sentido_pdf or "").strip().upper()
    if not s:
        return ""
    if "AMBOS" in s:
        if fam == "MG050":
            return "Oeste - SP/Leste - BH"
        return "Sul - SP/Norte - BH"
    # DECRESCENTE antes de CRESCENTE (evita match dentro de "DECRESCENTE")
    if "DECRESCENTE" in s:
        if fam == "MG050":
            return "Leste - BH"
        return "Sul - SP"
    if "CRESCENTE" in s:
        if fam == "MG050":
            return "Oeste - SP"
        return "Norte - BH"
    return (sentido_pdf or "").strip()
