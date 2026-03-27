"""
Teste local do endpoint POST /nc/analisar-pdf (ZIP com PDF + XLSX).

Uso:
  1. Em um terminal, na raiz do projeto, suba a API:
       cd c:\\GeradorARTESP
       python -m uvicorn render_api.app:app --reload --port 8000

  2. Em outro terminal, rode o teste (pode ser de qualquer pasta):
       python c:\\GeradorARTESP\\render_api\\scripts\\test_analisar_pdf_local.py c:\\caminho\\para\\seu.pdf

  Ou da raiz do projeto:
       cd c:\\GeradorARTESP
       python render_api/scripts/test_analisar_pdf_local.py caminho/para/seu.pdf

  O script envia o PDF para http://127.0.0.1:8000/nc/analisar-pdf, salva o ZIP
  em Relatorio_Analise_NCs_*.zip (na pasta atual) e extrai em test_analise_saida/.

  --lote 50  Artemig (Exportar Kcor + imagens/PDF na mesma pasta do ZIP que o relatório).
  --lote 13  ARTESP (padrão do script).

  --teste-local  Força seções de alerta no PDF mesmo com data da constatação ≠ hoje.
  Ou defina ARTESP_NC_TESTE_LOCAL=1 ao subir o uvicorn (vale para todas as chamadas).
"""
from __future__ import annotations

import argparse
import sys
import zipfile
from pathlib import Path

try:
    import httpx
except ImportError:
    print("Instale httpx: pip install httpx")
    sys.exit(1)


def main():
    parser = argparse.ArgumentParser(description="Testa POST /nc/analisar-pdf localmente (ZIP = PDF + XLSX)")
    parser.add_argument("pdf", type=Path, help="Caminho do PDF de NC Constatação de Rotina")
    parser.add_argument("--port", type=int, default=8000, help="Porta do uvicorn (default 8000)")
    parser.add_argument("--out-zip", type=Path, default=Path("Relatorio_Analise_NCs.zip"), help="Arquivo ZIP de saída")
    parser.add_argument("--extract-dir", type=Path, default=Path("test_analise_saida"), help="Pasta para extrair o ZIP")
    parser.add_argument(
        "--teste-local",
        action="store_true",
        help="Força alertas no PDF (gap/emerg.) mesmo se a data da constatação ≠ hoje",
    )
    parser.add_argument(
        "--lote",
        default="13",
        help="Lote do formulário (13/21/26 ARTESP; 50 Artemig). Default: 13",
    )
    args = parser.parse_args()

    pdf_path = args.pdf.resolve()
    if not pdf_path.is_file():
        print(f"Erro: PDF não encontrado: {pdf_path}")
        print("Use o caminho real do seu arquivo PDF de NC Constatação (ex.: C:\\Users\\seu_usuario\\Documentos\\constatacao.pdf)")
        sys.exit(1)

    url = f"http://127.0.0.1:{args.port}/nc/analisar-pdf"
    print(f"Enviando {pdf_path.name} para {url} ...")

    with open(pdf_path, "rb") as f:
        pdf_bytes = f.read()

    try:
        with httpx.Client(timeout=120.0) as client:
            data = {"limiar_km": "2.0", "lote": (args.lote or "13").strip()}
            if args.teste_local:
                data["teste_local"] = "1"
            r = client.post(
                url,
                files=[("pdfs", (pdf_path.name, pdf_bytes, "application/pdf"))],
                data=data,
            )
    except httpx.ConnectError as e:
        print(f"Erro: não foi possível conectar em {url}")
        print("  Suba a API antes: python -m uvicorn render_api.app:app --reload --port 8000")
        print(e)
        sys.exit(1)

    if r.status_code != 200:
        print(f"Erro HTTP {r.status_code}")
        print(r.text[:1500] if r.text else "(sem corpo)")
        sys.exit(1)

    zip_path = args.out_zip.resolve()
    zip_path.write_bytes(r.content)
    print(f"ZIP salvo: {zip_path} ({len(r.content)} bytes)")

    with zipfile.ZipFile(zip_path, "r") as zf:
        nomes = zf.namelist()
        print(f"Arquivos no ZIP: {nomes}")
        tem_pdf = any(n.lower().endswith(".pdf") for n in nomes)
        tem_xlsx = any(n.lower().endswith(".xlsx") for n in nomes)
        if not tem_pdf or not tem_xlsx:
            print("Aviso: esperado PDF e XLSX no ZIP.")
        args.extract_dir.mkdir(parents=True, exist_ok=True)
        zf.extractall(args.extract_dir)
    print(f"Conteúdo extraído em: {args.extract_dir.resolve()}")
    print("OK: abra o PDF e o XLSX na pasta de saída para validar.")


if __name__ == "__main__":
    main()
