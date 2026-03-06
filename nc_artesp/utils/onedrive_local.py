"""
nc_artesp/utils/onedrive_local.py
Compatibilidade OneDrive — cópia local antes de abrir com COM/xlwings.
No contexto web (Linux/Render) esta função é um no-op.
"""

from __future__ import annotations

from pathlib import Path
from typing import Callable


def processar_com_copia_local(caminho: Path, func: Callable, *args, **kwargs):
    """
    Em produção web: chama func(caminho, *args, **kwargs) diretamente.
    No desktop pode copiar para %TEMP% antes para evitar bloqueio OneDrive.
    """
    return func(caminho, *args, **kwargs)
