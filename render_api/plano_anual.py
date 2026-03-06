# -*- coding: utf-8 -*-
"""
Transição do Plano Anual e Faxina.
- Seleção do plano mestre por mes/ano (ex.: Dezembro usa ano corrente; Janeiro usa ano corrente).
- Faxina: após 20 de Janeiro, remove plano do ano retrasado (ex.: Jan/2026 remove 2024).
- Planos ficam em pasta dedicada (não são apagados pela auto-limpeza de 24h dos ZIPs).
"""
from __future__ import annotations

import logging
import os
from datetime import datetime
from pathlib import Path
from typing import Optional

logger = logging.getLogger(__name__)


def _base_data_dir() -> Path:
    """Pasta de dados (ex.: /data no Render). Env ARTESP_DATA_DIR tem prioridade."""
    env = (os.getenv("ARTESP_DATA_DIR") or "").strip()
    if env:
        return Path(env).resolve()
    return (Path("/data") if os.name != "nt" else Path(os.getcwd()) / "data").resolve()


# Pasta dos planos anuais (ex.: /data/anual). Não é varrida pela limpeza de 24h.
PLANO_ANUAL_DIR = _base_data_dir() / "anual"

# Nome do arquivo: ANUAL_2025.xlsx, ANUAL_2026.xlsx (ou PLANO_ANUAL_{ano}.xlsx)
NOME_PLANO_PADRAO = "ANUAL_{ano}.xlsx"


def caminho_plano_ano(ano: int) -> Path:
    """Retorna o caminho do arquivo do plano anual para o ano dado."""
    return PLANO_ANUAL_DIR / NOME_PLANO_PADRAO.format(ano=ano)


def selecionar_plano_mestre(mes_executado: int, ano_executado: int) -> Optional[Path]:
    """
    Retorna o caminho do plano mestre a usar para o mês/ano processado.

    Ponto crítico: ano_executado deve ser sempre o ano em que o serviço foi realizado,
    e não o ano em que o arquivo está sendo subido. Ex.: em 10/Jan/2026 ao processar
    o executado de Dezembro/2025, ano_executado=2025 → busca ANUAL_2025.xlsx.
    Isso garante o "overlap" (sobreposição) no fluxo de transição de ano.
    """
    ano_busca = ano_executado
    path = caminho_plano_ano(ano_busca)
    if path.is_file():
        return path
    # Fallback: nome alternativo PLANO_ANUAL_{ano}.xlsx
    path_alt = PLANO_ANUAL_DIR / f"PLANO_ANUAL_{ano_busca}.xlsx"
    if path_alt.is_file():
        return path_alt
    return None


def carregar_plano_mestre(mes_executado: int, ano_executado: int, sheet_name=None):
    """
    Carrega o DataFrame do plano anual para o mês/ano indicado.
    ano_executado = ano em que o serviço foi realizado (não o ano da data de upload).
    Levanta FileNotFoundError se o arquivo não existir (para uso em auditoria).

    sheet_name: aba do Excel. Se None, usa índice 0 (primeira aba) para evitar
    diferença de nomes entre anos; pode padronizar internamente como "PROGRAMACAO".
    """
    path = selecionar_plano_mestre(mes_executado, ano_executado)
    if path is None:
        raise FileNotFoundError(
            f"Plano Anual de {ano_executado} não encontrado em {PLANO_ANUAL_DIR}. "
            f"Arquivo esperado: {NOME_PLANO_PADRAO.format(ano=ano_executado)} ou PLANO_ANUAL_{ano_executado}.xlsx"
        )
    import pandas as pd
    # Carregar por índice (0) ou nome padronizado, para não depender do nome da aba entre anos
    if sheet_name is not None:
        return pd.read_excel(path, sheet_name=sheet_name)
    try:
        return pd.read_excel(path, sheet_name=0)
    except Exception:
        return pd.read_excel(path, sheet_name="PROGRAMACAO")


def faxina_anual_antigo(dia_corte_jan: int = 20) -> dict:
    """
    Remove o plano anual do ano retrasado, apenas após o dia indicado de Janeiro
    (ex.: após 20/Jan permite fechar Dezembro do ano anterior com o mestre correto).
    Em Jan/2026, remove ANUAL_2024.xlsx e PLANO_ANUAL_2024.xlsx (mantém 2025 e 2026).
    Retorna {"removidos": N, "erros": M, "arquivos": [...], "executado": bool}.
    """
    hoje = datetime.now()
    resultado = {"removidos": 0, "erros": 0, "arquivos": [], "executado": False}

    if hoje.month != 1 or hoje.day < dia_corte_jan:
        logger.info(
            "Faxina anual: não executada (hoje=%s; só roda em Janeiro após dia %d)",
            hoje.date(),
            dia_corte_jan,
        )
        return resultado

    ano_remover = hoje.year - 2  # Ex.: 2026 -> remove 2024
    resultado["executado"] = True

    if not PLANO_ANUAL_DIR.is_dir():
        return resultado

    for nome in (NOME_PLANO_PADRAO.format(ano=ano_remover), f"PLANO_ANUAL_{ano_remover}.xlsx"):
        path = PLANO_ANUAL_DIR / nome
        if not path.is_file():
            continue
        try:
            path.unlink()
            resultado["removidos"] += 1
            resultado["arquivos"].append(nome)
            logger.info("Faxina anual: removido %s (ano retrasado %d)", path.name, ano_remover)
        except OSError as e:
            resultado["erros"] += 1
            logger.warning("Faxina anual: falha ao remover %s: %s", path, e)

    return resultado
