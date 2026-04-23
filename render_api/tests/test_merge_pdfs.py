"""Fusão de PDFs (Constatacoes_unificadas) para /nc/analisar-pdf e /nc/extrair-pdf."""

import pytest

from nc_artesp.pdf_extractor import FITZ_OK, merge_pdfs_bytes


@pytest.mark.skipif(not FITZ_OK, reason="pymupdf não instalado")
def test_merge_pdfs_bytes_um_retorna_o_mesmo():
    import fitz

    d = fitz.open()
    d.new_page()
    b = d.tobytes()
    d.close()
    assert merge_pdfs_bytes([b]) == b


@pytest.mark.skipif(not FITZ_OK, reason="pymupdf não instalado")
def test_merge_pdfs_bytes_dois_documentos():
    import fitz

    def _um_pdf() -> bytes:
        d = fitz.open()
        d.new_page()
        out = d.tobytes()
        d.close()
        return out

    a, c = _um_pdf(), _um_pdf()
    m = merge_pdfs_bytes([a, c])
    doc = fitz.open(stream=m, filetype="pdf")
    try:
        assert doc.page_count == 2
    finally:
        doc.close()
