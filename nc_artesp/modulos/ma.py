"""
nc_artesp/modulos/ma.py
────────────────────────────────────────────────────────────────────────────
Módulo Meio Ambiente: funções e geração dos equivalentes aos Módulos 1, 2, 3 e 4
exclusivos para MA (Meio Ambiente).

  M01 MA: EAF desde PDF + Separar NC → Exportar MA (arquivos EAF individuais)
  M02 MA: Gerar Modelo Foto (Kria + Resposta) desde Exportar MA
  M03 MA: Inserir NC / Kcor-Kria Meio Ambiente (equivalente M07)
  M04 MA: Juntar Kcor-Kria MA → Acumulado MA

Uso: importar e chamar executar_m01_ma, executar_m02_ma, executar_m03_ma, executar_m04_ma
      ou executar_pipeline_ma_completo para rodar a sequência 1→2→3→4.
"""

from __future__ import annotations

import logging
from pathlib import Path
from typing import Callable

from config import (
    M01_MA_EAF,
    M01_MA_EXPORTAR,
    M02_MA_KRIA,
    M02_MA_PENDENTES,
    M04_MA_ACUMULADO,
    M04_MA_ENTRADA,
    M04_MA_NOME_SAIDA,
    M04_MA_SAIDA,
    M07_ENTRADA,
    M07_IMAGENS,
    M07_MODELO_KCOR,
    M07_SAIDA,
)
from utils.helpers import garantir_pasta, resolver_path_ficheiro_ci

logger = logging.getLogger(__name__)


def executar_m01_ma(
    pdf_bytes: bytes | None = None,
    eaf_planilha_mae: Path | None = None,
    pasta_saida_eaf: Path | None = None,
    pasta_exportar: Path | None = None,
    nome_arquivo_eaf: str = "EAF_MA_desde_PDF.xlsx",
    list_pdf_bytes: list[bytes] | None = None,
    gerar_separar_nc: bool = True,
    callback_progresso: Callable[[int, int, str], None] | None = None,
    sobrescrever: bool = True,
) -> dict:
    """
    Módulo 1 MA: gera a planilha EAF (planilha-mãe) e/ou executa Separar NC.

    Entrada:
      - pdf_bytes: um PDF de Meio Ambiente (gera EAF desde o texto do PDF).
      - list_pdf_bytes: lista de PDFs (gera EAF consolidado).
      - eaf_planilha_mae: path do EAF já existente (só Separar NC).

    Saída:
      - pasta_saida_eaf: onde gravar o EAF (default: M01_MA_EAF).
      - pasta_exportar: onde gravar os XLS individuais do Separar NC (default: M01_MA_EXPORTAR).

    Retorna dict com: eaf (Path | None), arquivos_separados (list[Path]).
    """
    from . import inserir_nc_kria
    from . import separar_nc

    pasta_saida_eaf = pasta_saida_eaf or M01_MA_EAF
    pasta_exportar = pasta_exportar or M01_MA_EXPORTAR
    garantir_pasta(pasta_saida_eaf)
    garantir_pasta(pasta_exportar)

    eaf_path = eaf_planilha_mae
    if eaf_path is None and (pdf_bytes is not None or list_pdf_bytes):
        if list_pdf_bytes:
            eaf_path = inserir_nc_kria.gerar_eaf_desde_pdfs_ma(
                list_pdf_bytes,
                pasta_saida=pasta_saida_eaf,
                nome_arquivo=nome_arquivo_eaf,
            )
        else:
            eaf_path = inserir_nc_kria.gerar_eaf_desde_pdf_ma(
                pdf_bytes,
                pasta_saida=pasta_saida_eaf,
                nome_arquivo=nome_arquivo_eaf,
            )
        if eaf_path:
            logger.info("M01 MA: EAF gerado: %s", eaf_path.name)

    arquivos_separados: list[Path] = []
    if gerar_separar_nc and eaf_path is not None and eaf_path.is_file():
        try:
            arquivos_separados = separar_nc.executar(
                arquivo_mae=eaf_path,
                pasta_destino=pasta_exportar,
                callback_progresso=callback_progresso,
                sobrescrever=sobrescrever,
            )
            logger.info("M01 MA: Separar NC gerou %s arquivo(s).", len(arquivos_separados))
        except Exception as e:
            logger.warning("M01 MA: Separar NC falhou: %s", e)

    return {"eaf": eaf_path, "arquivos_separados": arquivos_separados}


def executar_m02_ma(
    pasta_xls: Path | None = None,
    pasta_saida_kria: Path | None = None,
    pasta_saida_resp: Path | None = None,
    pasta_fotos_nc: Path | None = None,
    pasta_fotos_pdf: Path | None = None,
    modelo_kria: Path | None = None,
    modelo_resposta: Path | None = None,
    callback_progresso: Callable[[int, int, str], None] | None = None,
) -> dict:
    """
    Módulo 2 MA: gera Kria e Resposta (modelo foto) a partir dos XLS de Exportar MA.

    pasta_xls: pasta com os .xlsx individuais (saída do M01 MA). Default: M01_MA_EXPORTAR.
    pasta_saida_kria: default M02_MA_KRIA.
    pasta_saida_resp: default M02_MA_PENDENTES.

    Retorna dict com: kria (list[Path]), resposta (list[Path]), erros (list[str]).
    """
    from . import gerar_modelo_foto

    pasta_xls = pasta_xls or M01_MA_EXPORTAR
    pasta_saida_kria = pasta_saida_kria or M02_MA_KRIA
    pasta_saida_resp = pasta_saida_resp or M02_MA_PENDENTES
    pasta_fotos_nc = pasta_fotos_nc or M07_IMAGENS
    pasta_fotos_pdf = pasta_fotos_pdf or M07_IMAGENS
    garantir_pasta(pasta_saida_kria)
    garantir_pasta(pasta_saida_resp)

    resultado = gerar_modelo_foto.executar(
        pasta_xls=pasta_xls,
        modelo_kria=modelo_kria,
        pasta_saida_kria=pasta_saida_kria,
        modelo_resposta=modelo_resposta,
        pasta_saida_resp=pasta_saida_resp,
        pasta_fotos_nc=pasta_fotos_nc,
        pasta_fotos_pdf=pasta_fotos_pdf,
        callback_progresso=callback_progresso,
    )
    logger.info(
        "M02 MA: %s Kria, %s Resposta, %s erro(s).",
        len(resultado.get("kria", [])),
        len(resultado.get("resposta", [])),
        len(resultado.get("erros", [])),
    )
    return resultado


def executar_m03_ma(
    pdf_bytes: bytes | None = None,
    pasta_entrada: Path | None = None,
    pasta_imagens: Path | None = None,
    modelo_kcor: Path | None = None,
    pasta_saida: Path | None = None,
    nome_origem: str = "PDF MA",
) -> list[Path]:
    """
    Módulo 3 MA: Inserir NC Meio Ambiente → gera Kcor-Kria e imagens.

    Entrada:
      - pdf_bytes: processa um único PDF (extrai texto + gera Kcor-Kria).
      - pasta_entrada: se não passar PDF, processa os .xlsx da pasta (modo formulário).

    Saída em pasta_saida (default: M07_SAIDA).

    Retorna lista de Path dos Kcor-Kria gerados.
    """
    from . import inserir_nc_kria

    pasta_imagens = pasta_imagens or M07_IMAGENS
    modelo_kcor = resolver_path_ficheiro_ci(modelo_kcor or M07_MODELO_KCOR)
    pasta_saida = pasta_saida or M07_SAIDA

    if pdf_bytes is not None:
        destino = inserir_nc_kria._processar_pdf_meio_ambiente(
            pdf_bytes,
            pasta_imagens=pasta_imagens,
            modelo_kcor=modelo_kcor,
            pasta_saida=pasta_saida,
            nome_origem=nome_origem,
        )
        return [destino] if destino else []
    # Modo pasta (formulários Kria preenchidos)
    return inserir_nc_kria.executar_meio_ambiente(
        pasta_entrada=pasta_entrada or M07_ENTRADA,
        pasta_imagens=pasta_imagens,
        modelo_kcor=modelo_kcor,
        pasta_saida=pasta_saida,
    )


def executar_m04_ma(
    pasta_entrada: Path | None = None,
    arquivo_acumulado: Path | None = None,
    pasta_saida: Path | None = None,
    nome_saida: str | None = None,
    callback_progresso: Callable[[int, int, str], None] | None = None,
    arquivos_entrada: list[Path] | None = None,
) -> Path | None:
    """
    Módulo 4 MA: junta os Kcor-Kria de Meio Ambiente no acumulado MA.

    pasta_entrada: Kcor-Kria individuais (default: M04_MA_ENTRADA = Arquivos/Meio Ambiente).
    arquivo_acumulado: planilha acumulada existente (default: M04_MA_ACUMULADO).
    pasta_saida: onde salvar o relatório (default: M04_MA_SAIDA).
    nome_saida: nome do arquivo de saída (default: M04_MA_NOME_SAIDA).

    Retorna o Path do acumulado salvo ou None.
    """
    from . import juntar_arquivos

    pasta_entrada = pasta_entrada or M04_MA_ENTRADA
    arquivo_acumulado = arquivo_acumulado or M04_MA_ACUMULADO
    pasta_saida = pasta_saida or M04_MA_SAIDA
    nome_saida = nome_saida or M04_MA_NOME_SAIDA

    return juntar_arquivos.executar(
        pasta_entrada=pasta_entrada,
        arquivo_acumulado=arquivo_acumulado,
        pasta_saida=pasta_saida,
        nome_saida=nome_saida,
        callback_progresso=callback_progresso,
        arquivos_entrada=arquivos_entrada,
    )


def executar_pipeline_ma_completo(
    pdf_bytes: bytes | None = None,
    list_pdf_bytes: list[bytes] | None = None,
    pasta_base: Path | None = None,
    gerar_eaf: bool = True,
    gerar_separar_nc: bool = True,
    gerar_m02: bool = True,
    gerar_m03: bool = True,
    gerar_m04: bool = False,
    arquivo_acumulado_ma: Path | None = None,
    callback_progresso: Callable[[int, int, str], None] | None = None,
) -> dict:
    """
    Executa a sequência M01 MA → M02 MA → M03 MA e opcionalmente M04 MA.

    Se pdf_bytes ou list_pdf_bytes for informado:
      - M01 MA: gera EAF desde PDF e executa Separar NC.
      - M02 MA: gera Kria e Resposta desde Exportar MA.
      - M03 MA: gera Kcor-Kria desde o(s) PDF(s).
    Se gerar_m04=True, executa M04 MA (juntar); exige arquivo_acumulado_ma existente.

    pasta_base: se informado, subpastas EAF MA, Exportar MA, etc. são criadas sob ela.
    Retorna dict com: eaf, arquivos_separados, m02 (kria, resposta, erros), m03 (list[Path]), m04 (Path | None).
    """
    if pasta_base is not None:
        base = Path(pasta_base)
        pasta_eaf = base / "EAF MA"
        pasta_exportar = base / "Exportar MA"
        pasta_kria = base / "Arquivo Foto MA"
        pasta_resp = base / "Pendentes MA"
        pasta_imagens = base / "Imagens MA"
        pasta_kcor = base / "Meio Ambiente"
        pasta_acum = base / "Meio Ambiente" / "Acumulado"
    else:
        pasta_eaf = M01_MA_EAF
        pasta_exportar = M01_MA_EXPORTAR
        pasta_kria = M02_MA_KRIA
        pasta_resp = M02_MA_PENDENTES
        pasta_imagens = M07_IMAGENS
        pasta_kcor = M07_SAIDA
        pasta_acum = M04_MA_SAIDA

    resultado = {
        "eaf": None,
        "arquivos_separados": [],
        "m02": {"kria": [], "resposta": [], "erros": []},
        "m03": [],
        "m04": None,
    }

    pdf_list = list_pdf_bytes if list_pdf_bytes else ([pdf_bytes] if pdf_bytes else [])

    if gerar_eaf and pdf_list:
        r1 = executar_m01_ma(
            list_pdf_bytes=pdf_list if len(pdf_list) > 1 else None,
            pdf_bytes=pdf_list[0] if len(pdf_list) == 1 else None,
            pasta_saida_eaf=pasta_eaf,
            pasta_exportar=pasta_exportar,
            gerar_separar_nc=gerar_separar_nc,
            callback_progresso=callback_progresso,
        )
        resultado["eaf"] = r1.get("eaf")
        resultado["arquivos_separados"] = r1.get("arquivos_separados") or []

    if gerar_m02 and resultado["arquivos_separados"]:
        r2 = executar_m02_ma(
            pasta_xls=pasta_exportar,
            pasta_saida_kria=pasta_kria,
            pasta_saida_resp=pasta_resp,
            pasta_fotos_nc=pasta_imagens,
            pasta_fotos_pdf=pasta_imagens,
            callback_progresso=callback_progresso,
        )
        resultado["m02"] = {
            "kria": r2.get("kria") or [],
            "resposta": r2.get("resposta") or [],
            "erros": r2.get("erros") or [],
        }

    if gerar_m03 and pdf_list:
        for idx, pb in enumerate(pdf_list):
            nome = f"PDF MA {idx + 1}" if len(pdf_list) > 1 else "PDF MA"
            lista = executar_m03_ma(
                pdf_bytes=pb,
                pasta_imagens=pasta_imagens,
                pasta_saida=pasta_kcor,
                nome_origem=nome,
            )
            resultado["m03"].extend(lista)

    if gerar_m04:
        acu = arquivo_acumulado_ma or M04_MA_ACUMULADO
        if acu.is_file():
            resultado["m04"] = executar_m04_ma(
                pasta_entrada=pasta_kcor,
                arquivo_acumulado=acu,
                pasta_saida=pasta_acum,
                callback_progresso=callback_progresso,
            )
        else:
            logger.warning("M04 MA: arquivo acumulado não encontrado (%s). Ignorando M04.", acu)

    return resultado
