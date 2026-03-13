"""
render_api/fotos_router.py
────────────────────────────────────────────────────────────────────────────
Router FastAPI para o módulo de processamento de fotos de campo.

Fluxo: a listagem (Módulo 1) é criada a partir do ZIP de fotos.
A Relação Total (assets) fornece as coordenadas para preencher Rodovia/km.
O processamento organiza as fotos em subpastas por rodovia.

Endpoints:
  POST /fotos/listar             – ZIP de fotos → XLSX listagem (metadados + GPS)
  POST /fotos/coordenadas-km     – XLSX listagem + Relação Total → XLSX com Rodovia/km

Todos os endpoints:
  • Aceitam arquivos via multipart/form-data (UploadFile)
  • Retornam StreamingResponse com o arquivo gerado
  • Requerem autenticação JWT (Bearer token)
  • Limite de upload apenas local (versão web sem limite)

Desenvolvedor: Ozeias Engler
"""

from __future__ import annotations

import asyncio
import base64
import io
import json
import logging
import os
import re
import tempfile
import traceback
import unicodedata
import uuid
import zipfile
from datetime import datetime
from pathlib import Path
from typing import Optional

from fastapi import APIRouter, Depends, File, Form, HTTPException, Request, UploadFile
from fastapi.responses import FileResponse, StreamingResponse

logger = logging.getLogger(__name__)

# Limite de upload (MB); 0 = sem limite
def _ler_int_env(name: str, default: int) -> int:
    try:
        v = os.environ.get(name, "").strip()
        return int(v) if v else default
    except ValueError:
        return default


def _em_producao() -> bool:
    """True se em produção (Render)."""
    return bool(
        os.environ.get("RENDER") or
        os.environ.get("PRODUCTION") or
        os.environ.get("ARTESP_PRODUCTION")
    )


# Produção: default 512 MB (evita OOM). Local: 1024. 0 = sem limite.
_DEFAULT_UPLOAD_MB = 512 if _em_producao() else 1024
MAX_UPLOAD_MB = _ler_int_env("ARTESP_FOTOS_MAX_UPLOAD_MB", _DEFAULT_UPLOAD_MB)
MAX_UPLOAD_BYTES = MAX_UPLOAD_MB * 1024 * 1024 if MAX_UPLOAD_MB else 0


# ─── Autenticação (reutiliza o middleware do app principal) ───────────────────
def _get_auth(request: Request):
    """
    Verifica JWT Bearer ou cookie de sessão.
    Retorna payload ou lança HTTP 401.
    Importa a função do app principal em tempo de execução para evitar import circular.
    """
    try:
        from render_api.app import verificar_token_request  # type: ignore
        return verificar_token_request(request)
    except ImportError:
        try:
            from app import verificar_token_request           # type: ignore
            return verificar_token_request(request)
        except ImportError:
            return {}  # ambiente sem auth (desenvolvimento local)


# ─── Helpers ─────────────────────────────────────────────────────────────────
def _ler_upload(arquivo: UploadFile) -> bytes:
    """Lê UploadFile em bytes; aplica limite se MAX_UPLOAD_BYTES > 0."""
    conteudo = arquivo.file.read()
    if MAX_UPLOAD_BYTES > 0 and len(conteudo) > MAX_UPLOAD_BYTES:
        raise HTTPException(
            status_code=413,
            detail=f"Arquivo '{arquivo.filename}' excede o limite de {MAX_UPLOAD_MB} MB.",
        )
    return conteudo


def _stream(conteudo: bytes, filename: str, media_type: str) -> StreamingResponse:
    return StreamingResponse(
        io.BytesIO(conteudo),
        media_type=media_type,
        headers={"Content-Disposition": f'attachment; filename="{filename}"'},
    )


def _stream_zip(conteudo: bytes, filename: str) -> StreamingResponse:
    return _stream(conteudo, filename, "application/zip")


def _stream_xlsx(conteudo: bytes, filename: str) -> StreamingResponse:
    return _stream(
        conteudo, filename,
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


def _importar_core():
    """Importa core com mensagem de erro amigável se dependências faltarem."""
    try:
        from fotos_campo import core
        return core
    except ImportError as e:
        raise HTTPException(
            status_code=503,
            detail=(
                f"Módulo fotos_campo não disponível: {e}\n"
                "Verifique se openpyxl, pillow e piexif estão instalados."
            ),
        )


# ─── Assets: Relação Total de coordenadas por lote ───────────────────────────
# Cada módulo tem seu próprio assets: fotos_campo usa fotos_campo/assets
FOTOS_ASSETS = Path(__file__).resolve().parent.parent / "fotos_campo" / "assets"
# Template do relatório Foto 2 Lados (módulo nc_artesp)
NC_ARTESP_ASSETS = Path(__file__).resolve().parent.parent / "nc_artesp" / "assets"
FOTO2LADOS_MODELO_NOME = "Planilha Modelo Conservação - Foto 2 Lados.xlsx"
# Lote 13, 21 e 26 — arquivos em fotos_campo/assets (XLSX ou CSV)
RELACOES_POR_LOTE: dict[str, str] = {
    "13": "Relação Total - Lote 13.xlsx",
    "21": "Relação Total - Lote 21.xlsx",
    "26": "Relação Total - Lote 26.xlsx",
}
DEFAULT_LOTE = "13"

# Pasta em disco para downloads (compartilhada entre workers; evita "não encontrado")
_env_data = (os.environ.get("ARTESP_DATA_DIR") or "").strip()
_data_dir = Path(_env_data).resolve() if _env_data else (Path(__file__).resolve().parent.parent / "data")
FOTOS_DOWNLOADS_DIR = (_data_dir / "fotos_downloads").resolve()
FOTOS_DOWNLOAD_MAX_AGE_MINUTES = 15
try:
    FOTOS_DOWNLOADS_DIR.mkdir(parents=True, exist_ok=True)
except Exception as e:
    logger.warning("Pasta fotos_downloads não criada (%s), download usará só cache em memória: %s", FOTOS_DOWNLOADS_DIR, e)


def _limpar_downloads_antigos() -> None:
    """Remove arquivos em fotos_downloads mais antigos que FOTOS_DOWNLOAD_MAX_AGE_MINUTES."""
    if not FOTOS_DOWNLOADS_DIR.is_dir():
        return
    try:
        now = datetime.now()
        for p in FOTOS_DOWNLOADS_DIR.iterdir():
            if not p.is_file():
                continue
            try:
                mtime = datetime.fromtimestamp(p.stat().st_mtime)
                if (now - mtime).total_seconds() > FOTOS_DOWNLOAD_MAX_AGE_MINUTES * 60:
                    p.unlink(missing_ok=True)
            except Exception:
                pass
    except Exception as e:
        logger.debug("Limpeza de downloads antigos: %s", e)


# Tamanho máximo (bytes) para enviar o ZIP em base64 no SSE; acima disso usa GET /download
MAX_ZIP_INLINE_B64 = 100 * 1024 * 1024  # 100 MB — ZIP vem no SSE; evita GET que pode estourar timeout

# Cache em memória; limitado para evitar OOM (N entradas, B bytes)
_FOTOS_DOWNLOAD_CACHE: dict[str, tuple[bytes, str]] = {}
_FOTOS_CACHE_MAX_ENTRIES = 8
_FOTOS_CACHE_MAX_BYTES = 400 * 1024 * 1024


def _evictar_cache_fotos_ate_caber(zip_size: int) -> None:
    """Remove entradas antigas do cache até caber zip_size nos limites."""
    global _FOTOS_DOWNLOAD_CACHE
    total = sum(len(b) for b, _ in _FOTOS_DOWNLOAD_CACHE.values())
    keys_ordem = list(_FOTOS_DOWNLOAD_CACHE.keys())
    for kid in keys_ordem:
        if len(_FOTOS_DOWNLOAD_CACHE) <= _FOTOS_CACHE_MAX_ENTRIES and (total + zip_size) <= _FOTOS_CACHE_MAX_BYTES:
            break
        if kid not in _FOTOS_DOWNLOAD_CACHE:
            continue
        b, _ = _FOTOS_DOWNLOAD_CACHE.pop(kid, (b"", ""))
        total -= len(b)


def _sanitizar_rodovia(s: str) -> str:
    """Nome de pasta seguro para rodovia."""
    if not s or not isinstance(s, str):
        return "Sem_rodovia"
    s = re.sub(r'[\\/:*?"<>|]', "_", s.strip())
    return s[:80] if s else "Sem_rodovia"


# ─── Router ──────────────────────────────────────────────────────────────────
router = APIRouter(prefix="/fotos", tags=["Fotos de Campo"])


# ──────────────────────────────────────────────────────────────────────────────
#  MÓDULO 1 — Listar fotos + GPS EXIF
# ──────────────────────────────────────────────────────────────────────────────

@router.post(
    "/listar",
    summary="Gerar listagem de fotos (ZIP → XLSX)",
    response_description="XLSX com listagem gerada (metadados + GPS EXIF)",
)
async def listar_fotos(
    request: Request,
    fotos_zip: UploadFile = File(
        ...,
        description="ZIP contendo as fotos de campo (JPG/PNG). "
                    "Subpastas são suportadas."
    ),
):
    """
    Gera a listagem a partir do ZIP de fotos (não usa template em assets).
    Retorna um **XLSX** com: Caminho, Nome, Tipo, Tamanho, Lat/Lon (EXIF) e colunas
    em branco para Rodovia/km/Sentido (preenchidas pelo Módulo 2 com a Relação Total).
    """
    _get_auth(request)
    core = _importar_core()
    try:
        zip_bytes = _ler_upload(fotos_zip)
        xlsx_bytes, total = core.listar_de_zip(zip_bytes)
        return _stream_xlsx(xlsx_bytes, "fotos_listagem.xlsx")
    except HTTPException:
        raise
    except Exception as e:
        logger.error("listar-fotos: %s", traceback.format_exc())
        raise HTTPException(status_code=500, detail=str(e))


# ──────────────────────────────────────────────────────────────────────────────
#  MÓDULO 2 — Coordenadas → Rodovia / km / Sentido
# ──────────────────────────────────────────────────────────────────────────────

@router.post(
    "/coordenadas-km",
    summary="Preencher Rodovia/km/Sentido por coordenada GPS",
    response_description="XLSX com colunas G,H,I,J preenchidas",
)
async def coordenadas_km(
    request: Request,
    listagem: UploadFile = File(
        ...,
        description="XLSX gerado pelo Módulo 1 (fotos_listagem.xlsx)"
    ),
    relacao_total: UploadFile = File(
        ...,
        description="XLSX da Relação Total do lote "
                    "(A=Rodovia, B=km, C=Sentido, D=Lat, E=Lon)"
    ),
):
    """
    Cruza as coordenadas GPS de cada foto (colunas E e F do XLSX de listagem)
    com os pontos da **Relação Total** e preenche:
    - G: Rodovia
    - H: km
    - I: Sentido
    - J: Distância ao ponto mais próximo (km)
    """
    _get_auth(request)
    core = _importar_core()
    try:
        xlsx_dados  = _ler_upload(listagem)
        xlsx_relac  = _ler_upload(relacao_total)
        xlsx_result = core.coordenadas_km_bytes(xlsx_dados, xlsx_relac)
        return _stream_xlsx(xlsx_result, "fotos_listagem_km.xlsx")
    except HTTPException:
        raise
    except Exception as e:
        logger.error("coordenadas-km: %s", traceback.format_exc())
        raise HTTPException(status_code=500, detail=str(e))


# ──────────────────────────────────────────────────────────────────────────────
#  PROCESSAR COMPLETO — Listar + Coordenadas + ZIP por rodovia
# ──────────────────────────────────────────────────────────────────────────────

def _csv_relacao_para_xlsx_bytes(csv_path: Path) -> bytes:
    """Lê CSV da Relação Total (; separador, header Rodovia;Km;Sentido;Latitude;Longitude) e retorna XLSX em bytes."""
    import csv
    wb = __import__("openpyxl").Workbook()
    ws = wb.active
    with open(csv_path, newline="", encoding="utf-8-sig") as f:
        reader = csv.reader(f, delimiter=";")
        for row_idx, row in enumerate(reader, 1):
            for col_idx, val in enumerate(row[:5], 1):  # A–E
                ws.cell(row=row_idx, column=col_idx, value=val.strip() if isinstance(val, str) else val)
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


def _normalizar_nome(s: str) -> str:
    """Normaliza nome de arquivo para comparação (NFC), evitando diferença ç vs c+cedilha."""
    return unicodedata.normalize("NFC", s)


def _carregar_relacao_assets(lote: str = DEFAULT_LOTE) -> bytes:
    """Carrega a Relação Total do lote (13, 21 ou 26) a partir de fotos_campo/assets. Retorna XLSX em bytes."""
    nome = RELACOES_POR_LOTE.get(lote) or RELACOES_POR_LOTE.get(DEFAULT_LOTE)
    if not nome:
        nome = RELACOES_POR_LOTE[DEFAULT_LOTE]
    path = FOTOS_ASSETS / nome
    if not path.is_file():
        nome_nfc = _normalizar_nome(nome)
        base = Path(nome).stem
        path_csv = FOTOS_ASSETS / (base + ".csv")
        if path_csv.is_file():
            path = path_csv
        elif FOTOS_ASSETS.is_dir():
            for p in FOTOS_ASSETS.iterdir():
                if p.is_file() and _normalizar_nome(p.name) == nome_nfc:
                    path = p
                    break
            else:
                path = None
        else:
            path = None
        if path is None or not path.is_file():
            try:
                existentes = list(FOTOS_ASSETS.iterdir()) if FOTOS_ASSETS.is_dir() else []
                lista = ", ".join(p.name for p in existentes[:15] if p.is_file()) or "(pasta vazia ou inexistente)"
            except Exception:
                lista = "(não foi possível listar)"
            raise HTTPException(
                status_code=503,
                detail=(
                    f"Relação Total do lote {lote} não encontrada. "
                    f"Pasta usada: fotos_campo/assets (assets do módulo). "
                    f"Arquivo esperado: {nome} (ou {base}.csv). "
                    f"Arquivos na pasta: {lista}"
                ),
            )
    if path.suffix.lower() == ".csv":
        return _csv_relacao_para_xlsx_bytes(path)
    return path.read_bytes()


def _carregar_modelo_foto2lados() -> Optional[bytes]:
    """Carrega o template do relatório Foto 2 Lados de nc_artesp/assets. Retorna None se não existir."""
    nome = FOTO2LADOS_MODELO_NOME
    path = NC_ARTESP_ASSETS / nome
    if path.is_file():
        return path.read_bytes()
    if NC_ARTESP_ASSETS.is_dir():
        nome_nfc = _normalizar_nome(nome)
        for p in NC_ARTESP_ASSETS.iterdir():
            if p.is_file() and _normalizar_nome(p.name) == nome_nfc:
                return p.read_bytes()
    return None


@router.post(
    "/processar-completo",
    summary="Processar fotos e devolver listagem + fotos por rodovia em um ZIP",
    response_description="ZIP com Listagem/ e Por_rodovia/",
)
async def processar_completo(
    request: Request,
    fotos_zip: UploadFile = File(
        ...,
        description="ZIP com as fotos de campo (JPG/PNG).",
    ),
    lote: str = Form("13", description="Lote da Relação Total: 13, 21 ou 26"),
):
    """
    Executa automaticamente: gera a listagem a partir do ZIP →
    preenche Rodovia/km com a Relação Total do lote em assets →
    organiza as fotos em subpastas por rodovia.
    Devolve um único ZIP com: Listagem/, Por_rodovia/.
    """
    _get_auth(request)
    core = _importar_core()
    lote_ok = (lote or "").strip() or DEFAULT_LOTE
    if lote_ok not in RELACOES_POR_LOTE:
        lote_ok = DEFAULT_LOTE
    zip_bytes     = _ler_upload(fotos_zip)
    relacao_bytes = _carregar_relacao_assets(lote_ok)

    xlsx_listagem, _ = core.listar_de_zip(zip_bytes)
    xlsx_km          = core.coordenadas_km_bytes(xlsx_listagem, relacao_bytes)
    zip_foto2lados   = None
    modelo_2lados    = _carregar_modelo_foto2lados()
    if modelo_2lados:
        try:
            assunto = f"Fotos de Campo Lote {lote_ok}"
            zip_foto2lados = core.relatorio_foto2lados_bytes(xlsx_km, modelo_2lados, zip_bytes, assunto)
        except Exception as e:
            logger.warning("Relatório Foto 2 Lados não gerado: %s", e)
    zip_final = _montar_zip_fotos(xlsx_listagem, xlsx_km, zip_bytes, lote=lote_ok, zip_foto2lados=zip_foto2lados)
    nome_zip = f"Fotos_Processamento_Lote{lote_ok}_{datetime.now().strftime('%Y%m%d_%H%M')}.zip"
    return _stream_zip(zip_final, nome_zip)


def _montar_zip_fotos(
    xlsx_listagem: bytes,
    xlsx_km: bytes,
    zip_bytes: bytes,
    lote: str = "",
    zip_foto2lados: Optional[bytes] = None,
) -> bytes:
    """Monta o ZIP final: Listagem/ (planilhas + lote + relatório Foto 2 Lados) + Por_rodovia/ (fotos)."""
    core = _importar_core()
    pares = core.listar_rodovia_por_caminho(xlsx_km)
    path_to_rodovia: dict[str, str] = {}
    for caminho_n, rodovia_s in pares:
        r = _sanitizar_rodovia(rodovia_s)
        path_to_rodovia[caminho_n]                       = r
        path_to_rodovia[caminho_n.replace("\\", "/")]    = r
        path_to_rodovia[Path(caminho_n).name]            = r

    buf  = io.BytesIO()
    seen: set[str] = set()

    def nome_unico(arc_dest: str) -> str:
        out = arc_dest
        n   = 1
        while out in seen:
            stem = arc_dest.rsplit(".", 1)[0] if "." in arc_dest else arc_dest
            ext  = "." + arc_dest.rsplit(".", 1)[1] if "." in arc_dest else ""
            out  = f"{stem}_{n}{ext}"
            n   += 1
        seen.add(out)
        return out

    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED, compresslevel=6) as zf_out:
        zf_out.writestr("Listagem/fotos_listagem.xlsx",    xlsx_listagem)
        zf_out.writestr("Listagem/fotos_listagem_km.xlsx", xlsx_km)
        if lote:
            zf_out.writestr(
                "Listagem/lote_utilizado.txt",
                f"Relação Total utilizada: Lote {lote}\nArquivo: {RELACOES_POR_LOTE.get(lote, '')}\n".encode("utf-8"),
            )
        if zip_foto2lados:
            try:
                with zipfile.ZipFile(io.BytesIO(zip_foto2lados), "r") as zf_2lados:
                    for info in zf_2lados.infolist():
                        if info.is_dir():
                            continue
                        conteudo = zf_2lados.read(info.filename)
                        nome_arq = Path(info.filename).name
                        if nome_arq.lower().endswith(".xlsx"):
                            if "excluida" in nome_arq.lower():
                                zf_out.writestr("Listagem/Excluidas_Foto_2_Lados.xlsx", conteudo)
                            else:
                                zf_out.writestr("Listagem/Relatorio_Foto_2_Lados.xlsx", conteudo)
            except Exception as e:
                logger.warning("Não foi possível incluir relatório Foto 2 Lados no ZIP: %s", e)

        with zipfile.ZipFile(io.BytesIO(zip_bytes), "r") as zf_orig:
            for info in zf_orig.infolist():
                if info.is_dir():
                    continue
                arc    = info.filename.replace("\\", "/").lstrip("/")
                rodov  = path_to_rodovia.get(arc) or path_to_rodovia.get(Path(arc).name) or "Sem_rodovia"
                pasta  = "Sem_coordenadas" if rodov == "Sem_rodovia" else rodov
                dest   = nome_unico(f"Por_rodovia/{pasta}/{Path(arc).name}")
                try:
                    zf_out.writestr(dest, zf_orig.read(info.filename))
                except Exception as e:
                    logger.warning("ZIP fotos: ignorando %s: %s", info.filename, e)
    return buf.getvalue()


@router.post(
    "/processar-completo-progresso",
    summary="Processar fotos com log em tempo real (SSE)",
    response_description="Stream de eventos com progresso e download_id ao final",
)
async def processar_completo_progresso(
    request: Request,
    fotos_zip: UploadFile = File(..., description="ZIP com as fotos de campo."),
    lote: str = Form("13", description="Lote da Relação Total: 13, 21 ou 26"),
):
    """Listagem + Rodovia/km + organização por rodovia com log em tempo real (SSE).
    Ao final envia download_id para GET /fotos/download/{id}."""
    _get_auth(request)
    lote_ok = (lote or "").strip() or DEFAULT_LOTE
    if lote_ok not in RELACOES_POR_LOTE:
        lote_ok = DEFAULT_LOTE
    zip_bytes     = _ler_upload(fotos_zip)
    relacao_bytes = _carregar_relacao_assets(lote_ok)
    core          = _importar_core()

    def _evt(obj):
        return f"data: {json.dumps(obj, ensure_ascii=False)}\n\n"

    async def event_generator():
        try:
            yield _evt({"type": "progress", "status": "Recebendo arquivo ZIP...", "progresso": 5})
            await asyncio.sleep(0.02)

            yield _evt({"type": "progress", "status": "Listando fotos e extraindo GPS EXIF...", "progresso": 10})
            xlsx_listagem, total_fotos = await asyncio.to_thread(core.listar_de_zip, zip_bytes)
            yield _evt({"type": "progress", "status": f"Listagem concluída: {total_fotos} foto(s).", "progresso": 40})
            await asyncio.sleep(0.02)

            yield _evt({"type": "progress", "status": "Preenchendo Rodovia/km (Relação em assets)...", "progresso": 50})
            xlsx_km = await asyncio.to_thread(core.coordenadas_km_bytes, xlsx_listagem, relacao_bytes)
            yield _evt({"type": "progress", "status": "Rodovia/km preenchidos.", "progresso": 65})
            await asyncio.sleep(0.02)

            zip_foto2lados = None
            modelo_2lados = _carregar_modelo_foto2lados()
            if modelo_2lados:
                yield _evt({"type": "progress", "status": "Gerando relatório Foto 2 Lados...", "progresso": 72})
                try:
                    assunto = f"Fotos de Campo Lote {lote_ok}"
                    zip_foto2lados = await asyncio.to_thread(
                        core.relatorio_foto2lados_bytes,
                        xlsx_km, modelo_2lados, zip_bytes, assunto,
                    )
                except Exception as e:
                    logger.warning("Relatório Foto 2 Lados não gerado: %s", e)

            yield _evt({"type": "progress", "status": "Organizando pastas por rodovia e montando ZIP...", "progresso": 80})
            zip_final = await asyncio.to_thread(
                _montar_zip_fotos,
                xlsx_listagem, xlsx_km, zip_bytes,
                lote_ok,
                zip_foto2lados,
            )
            yield _evt({"type": "progress", "status": "ZIP montado.", "progresso": 90})
            await asyncio.sleep(0.02)

            nome_zip = f"Fotos_Processamento_Lote{lote_ok}_{datetime.now().strftime('%Y%m%d_%H%M')}.zip"
            download_id = str(uuid.uuid4())
            _evictar_cache_fotos_ate_caber(len(zip_final))
            _FOTOS_DOWNLOAD_CACHE[download_id] = (zip_final, nome_zip)
            # Salvar também em disco para o download funcionar em qualquer worker/requisição
            try:
                FOTOS_DOWNLOADS_DIR.mkdir(parents=True, exist_ok=True)
                _limpar_downloads_antigos()
                zip_path = FOTOS_DOWNLOADS_DIR / f"{download_id}.zip"
                nome_path = FOTOS_DOWNLOADS_DIR / f"{download_id}.nome"
                zip_path.write_bytes(zip_final)
                nome_path.write_text(nome_zip, encoding="utf-8")
                logger.info("ZIP salvo em disco: %s", zip_path)
            except Exception as e:
                logger.warning("ZIP não salvo em disco (%s), download usará cache: %s", FOTOS_DOWNLOADS_DIR, e)
            # Para ZIPs pequenos, enviar em base64 no evento para o frontend baixar sem segunda requisição
            payload = {
                "type": "result",
                "status": "sucesso",
                "progresso": 100,
                "download_id": download_id,
                "nome_zip": nome_zip,
                "total_fotos": total_fotos,
            }
            if len(zip_final) <= MAX_ZIP_INLINE_B64:
                payload["zip_base64"] = base64.standard_b64encode(zip_final).decode("ascii")
            yield _evt(payload)
        except HTTPException:
            raise
        except Exception as e:
            logger.exception("processar-completo-progresso")
            yield _evt({"type": "result", "status": "error", "detail": str(e), "progresso": 0})

    return StreamingResponse(
        event_generator(),
        media_type="text/event-stream",
        headers={"Cache-Control": "no-cache", "X-Accel-Buffering": "no"},
    )


def _safe_filename_for_header(nome: str) -> str:
    """Nome seguro para Content-Disposition (evita quebrar o header)."""
    if not nome or not isinstance(nome, str):
        return "Fotos_Processamento.zip"
    # Remove caracteres que quebram o header
    safe = re.sub(r'[\\\r\n"]', "_", nome.strip())[:200]
    return safe or "Fotos_Processamento.zip"


@router.get(
    "/download/{download_id}",
    summary="Baixar ZIP gerado (após processar-completo-progresso)",
)
async def fotos_download(request: Request, download_id: str):
    """Retorna o ZIP (disco ou cache em memória). Requer autenticação."""
    _get_auth(request)
    if not download_id or re.search(r"[^a-zA-Z0-9\-]", download_id):
        raise HTTPException(status_code=404, detail="Download não encontrado ou expirado.")
    # 1) Tentar disco (funciona com múltiplos workers e entre requisições)
    zip_path = FOTOS_DOWNLOADS_DIR / f"{download_id}.zip"
    nome_path = FOTOS_DOWNLOADS_DIR / f"{download_id}.nome"
    if zip_path.is_file():
        try:
            nome_zip = nome_path.read_text(encoding="utf-8").strip() if nome_path.is_file() else "Fotos_Processamento.zip"
            filename_safe = _safe_filename_for_header(nome_zip)
            # FileResponse faz streaming direto do disco: cliente começa a receber logo, sem carregar tudo na memória
            return FileResponse(
                path=str(zip_path),
                media_type="application/zip",
                filename=filename_safe,
            )
        except Exception as e:
            logger.warning("Erro ao ler ZIP do disco: %s", e)
    # 2) Fallback: cache em memória (mesmo processo)
    if download_id in _FOTOS_DOWNLOAD_CACHE:
        zip_bytes, nome_zip = _FOTOS_DOWNLOAD_CACHE.pop(download_id)
        filename_safe = _safe_filename_for_header(nome_zip)
        return StreamingResponse(
            io.BytesIO(zip_bytes),
            media_type="application/zip",
            headers={
                "Content-Disposition": f'attachment; filename="{filename_safe}"',
                "Content-Length": str(len(zip_bytes)),
            },
        )
    # Diagnóstico: log quando o download não for encontrado
    _dir = FOTOS_DOWNLOADS_DIR.resolve()
    logger.warning(
        "Download não encontrado: id=%s, pasta=%s, existe=%s, arquivos=%s",
        download_id, _dir, _dir.exists(), list(_dir.glob("*.zip"))[:5] if _dir.exists() else []
    )
    raise HTTPException(status_code=404, detail="Download não encontrado ou expirado. Processe as fotos novamente.")


# ──────────────────────────────────────────────────────────────────────────────
#  INFO — Lista de endpoints disponíveis
# ──────────────────────────────────────────────────────────────────────────────

@router.get("/", summary="Status do módulo Fotos de Campo")
async def fotos_info():
    try:
        import openpyxl
        openpyxl_ok = True
    except ImportError:
        openpyxl_ok = False
    try:
        from PIL import Image
        pil_ok = True
    except ImportError:
        pil_ok = False
    try:
        import piexif
        piexif_ok = True
    except ImportError:
        piexif_ok = False
    try:
        import fitz
        fitz_ok = True
    except ImportError:
        fitz_ok = False

    # Lotes com Relação Total disponível em assets
    lotes_disponiveis = [
        {"lote": k, "arquivo": v, "disponivel": (FOTOS_ASSETS / v).is_file()}
        for k, v in RELACOES_POR_LOTE.items()
    ]
    return {
        "modulo": "fotos_campo",
        "versao": "1.0.0",
        "dependencias": {
            "openpyxl": openpyxl_ok,
            "pillow":   pil_ok,
            "piexif":   piexif_ok,
            "pymupdf":  fitz_ok,
        },
        "relacoes_lote": lotes_disponiveis,
        "endpoints": [
            "POST /fotos/listar                    → XLSX listagem gerada (metadados + GPS)",
            "POST /fotos/coordenadas-km            → XLSX com Rodovia/km/Sentido",
            "POST /fotos/processar-completo        → ZIP com Listagem/ + Por_rodovia/ (param: lote=13|21|26)",
            "POST /fotos/processar-completo-progresso → SSE com log; GET /fotos/download/{id} (param: lote=13|21|26)",
        ],
    }
