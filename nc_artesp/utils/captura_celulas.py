"""
nc_artesp/utils/captura_celulas.py
Captura de range de células como imagem (xlwings/COM — só Windows com Excel).
Em ambiente web retorna None sem quebrar o import.
"""

from __future__ import annotations

from typing import Optional, Union
from pathlib import Path

from .helpers import str_caminho_io_windows

# Tamanhos (px) dos recortes de imagem por modo (conservação / meio ambiente)
TAMANHO_CONSERVACAO = (275, 210)   # largura × altura em pixels
TAMANHO_MA          = (275, 210)


def _col_num_to_letter(n: int) -> str:
    """Converte número de coluna 1-based em letra(s) Excel (1=A, 26=Z, 27=AA)."""
    s = ""
    while n > 0:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s


def _range_from_bounds(row_ini: int, col_ini: int, row_fim: int, col_fim: int) -> str:
    """Monta string de range Excel a partir de linhas/colunas 1-based."""
    return f"{_col_num_to_letter(col_ini)}{row_ini}:{_col_num_to_letter(col_fim)}{row_fim}"


def exportar_range_como_imagem(
    arquivo: Optional[Path] = None,
    sheet_name: Union[int, str, None] = None,
    range_str: Optional[str] = None,
    destino: Optional[Path] = None,
    tamanho: tuple = TAMANHO_CONSERVACAO,
    *,
    wb_path: Optional[Path] = None,
    sheet_index: Optional[int] = None,
    row_ini: Optional[int] = None,
    col_ini: Optional[int] = None,
    row_fim: Optional[int] = None,
    col_fim: Optional[int] = None,
    largura: Optional[int] = None,
    altura: Optional[int] = None,
    forcar_fallback: bool = False,
) -> Optional[Path]:
    """
    Exporta um range de células como JPG (xlwings/COM — Windows + Excel).

    Aceita dois estilos de chamada:

    1) (arquivo, sheet_name, range_str, destino, tamanho)
       Ex.: exportar_range_como_imagem(path, "Plan1", "C6:F10", out, (275, 210))

    2) Parâmetros nomeados do inserir_nc_kria:
       wb_path, sheet_index, row_ini, col_ini, row_fim, col_fim, destino, largura, altura
       (forcar_fallback é ignorado; xlwings é usado quando disponível)
    """
    # Normalizar chamada no estilo inserir_nc_kria (wb_path + row/col)
    if wb_path is not None and row_ini is not None and col_ini is not None and row_fim is not None and col_fim is not None and destino is not None:
        arquivo = wb_path
        range_str = _range_from_bounds(row_ini, col_ini, row_fim, col_fim)
        sheet_name = sheet_index if sheet_index is not None else 0
        if largura is not None and altura is not None:
            tamanho = (largura, altura)

    if arquivo is None or range_str is None or destino is None:
        return None

    if sheet_name is None:
        sheet_name = 0

    try:
        import xlwings as xw
        with xw.App(visible=False, add_book=False) as app:
            wb = app.books.open(str_caminho_io_windows(arquivo))
            try:
                ws = wb.sheets[sheet_name]
                rng = ws.range(range_str)
                dest_io = str_caminho_io_windows(destino)
                rng.to_png(dest_io)
                if Path(dest_io).exists():
                    # Redimensionar se necessário
                    try:
                        from PIL import Image
                        img = Image.open(dest_io)
                        img = img.resize(tamanho, Image.LANCZOS)
                        img.save(dest_io, "JPEG", quality=90)
                    except Exception:
                        pass
                    return destino
            finally:
                wb.close()
    except Exception:
        pass
    return None
