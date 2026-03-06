"""
nc_artesp/utils/image_anchor.py
Utilitários para inserção de imagens ancoradas em células openpyxl.
"""

from __future__ import annotations


def _find_merged_range(ws, col: int, row: int):
    """
    Retorna (col_fim, row_fim) do merged range que contém (col, row).
    Se a célula não estiver em merge, retorna (col, row).
    """
    for mr in ws.merged_cells.ranges:
        if mr.min_col <= col <= mr.max_col and mr.min_row <= row <= mr.max_row:
            return mr.max_col, mr.max_row
    return col, row


def get_merged_bounds(ws, col: int, row: int) -> tuple[int, int, int, int]:
    """
    Retorna (min_col, min_row, max_col, max_row) do merged range que contém (col, row).
    Se a célula não estiver em merge, retorna (col, row, col, row).
    """
    for mr in ws.merged_cells.ranges:
        if mr.min_col <= col <= mr.max_col and mr.min_row <= row <= mr.max_row:
            return mr.min_col, mr.min_row, mr.max_col, mr.max_row
    return col, row, col, row


def patch_add_image(ws) -> None:
    """
    Monkey-patch em ws.add_image para garantir compatibilidade com versões
    do openpyxl que exigem anchor explícito.
    Na prática, delega ao add_image original — o módulo já monta o anchor
    internamente via _ImageFromBytes e TwoCellAnchor.
    """
    # No-op: o módulo gerar_modelo_foto.py constrói o anchor manualmente
    # antes de chamar ws.add_image, por isso não precisamos sobrescrever.
    pass
