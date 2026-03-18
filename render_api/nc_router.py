"""
render_api/nc_router.py
────────────────────────────────────────────────────────────────────────────
Router FastAPI para o pipeline de Não Conformidades (NC Artesp).

Modos de uso:
  • Etapa isolada (single-shot): sem job_id → cada endpoint cria job novo, processa,
    grava em final/, marca finished (retain 24h), retorna job_id e download_urls.
  • Pipeline completo: primeira chamada sem job_id cria o job; seguintes enviam job_id
    (Form), reutilizam o mesmo workspace (stage1/, stage2/, final/). Só marca finished
    quando created ou finalize=1 (retain 72h); senão status=running e stage=stage1|stage2.

Ordem típica do pipeline: separar → gerar-modelo-foto → inserir-conservacao (ou
inserir-meio-ambiente) → juntar → inserir-numero. Alternativa em 2 chamadas:
  POST /nc/start (1 EAF) → POST /nc/stage2 (job_id + params).

Endpoints:
  POST /nc/extrair-pdf            – PDF NC Constatação → ZIP com nc(N).jpg e PDF(N).jpg
  POST /nc/analisar-pdf           – PDF NC Constatação → ZIP (PDF de análise + XLSX template preenchido)
  POST /nc/separar                → M01: EAF Excel → ZIP com XLS individuais
  POST /nc/gerar-modelo-foto      → M02: XLS ZIP + modelos → ZIP Kria + Resposta
  POST /nc/inserir-conservacao    → M03: Kria ZIP → ZIP Kcor-Kria Conservação
  POST /nc/inserir-meio-ambiente  → M07: Kria MA ZIP → ZIP Kcor-Kria MA
  POST /nc/juntar                 → M04: Kcor ZIP → XLSX acumulado
  POST /nc/inserir-numero         → M05: acumulado + nº inicial → XLSX numerado
  POST /nc/exportar-calendario    → M06: acumulado → arquivo .ics (iCalendar)
  GET  /nc/                       → status e deps
  (M08 Organizar Imagens existia apenas nas macros VBA — removido do fluxo web.)
"""

from __future__ import annotations

import asyncio
import io
import json
import logging
import os
import platform
import re
import secrets
import shutil
import sys
import tempfile
import time
import traceback
import unicodedata
import zipfile
from dataclasses import dataclass
from datetime import datetime, timedelta, timezone
from pathlib import Path
from typing import Any, Dict, List, Optional

from fastapi import APIRouter, Body, File, Form, HTTPException, Request, UploadFile
from fastapi.responses import JSONResponse, Response, StreamingResponse
from pydantic import BaseModel

try:
    from render_api.job_manager import carregar_job_nc as job_manager_carregar
except ImportError:
    job_manager_carregar = None

logger = logging.getLogger(__name__)

MAX_MB = 200
MAX_BYTES = MAX_MB * 1024 * 1024

# Tenta, em ordem:
#   1. Variável de ambiente ARTESP_NC_PROJ (configurável no Render)
#   2. Pasta nc_artesp/ dentro do próprio repositório (deploy Render)
#   3. Caminho Windows local (desenvolvimento desktop)
def _resolver_nc_proj() -> Path:
    # 1. Env var (mais flexível — configure no painel do Render se necessário)
    env = (
        __import__("os").getenv("ARTESP_NC_PROJ") or ""
    ).strip()
    if env:
        p = Path(env)
        if p.exists():
            return p

    # 2. Dentro do repo: GeradorARTESP/nc_artesp/
    repo_path = Path(__file__).resolve().parent.parent / "nc_artesp"
    if repo_path.exists():
        return repo_path

    # 3. Fallback local Windows (só funciona no desktop do dev)
    win_path = Path(r"C:\AUTOMAÇÃO_MACROS\Macros Kcor Ellen\artesp_nc_v2.0")
    return win_path  # pode não existir no Render — _nc_proj_disponivel() verifica


_NC_PROJ   = _resolver_nc_proj()
# Raiz do repositório (para importar nc_artesp como pacote)
_REPO_ROOT = Path(__file__).resolve().parent.parent

# Sempre resolve pelo repositório, independente do caminho do projeto local.
_NC_ASSETS = Path(__file__).resolve().parent.parent / "nc_artesp" / "assets" / "templates"
_NC_ARTEMIG_TEMPLATES = Path(__file__).resolve().parent.parent / "nc_artemig" / "assets" / "Template"

# ── Nomes de pastas dos relatórios (alinhados às macros nc_artesp/config.py) ───
DIR_EXPORTAR = "Exportar"
DIR_IMAGENS_PDF = "Imagens Provisórias - PDF"
DIR_KRIA = "Kria"
DIR_RESPOSTAS_PENDENTES = "Respostas Pendentes"
DIR_IMAGENS_CONSERVACAO = "Imagens Conservação"
DIR_CONSERVACAO = "Conservação"
DIR_IMAGENS_MA = "Imagens Meio Ambiente"
DIR_MA = "Meio Ambiente"
DIR_ACUMULADO = "Acumulado"
DIR_KCOR_CONSERVACAO = "Kcor Conservação"
DIR_EXPORTAR_EAF = "Exportar_EAF"

# Estratégias: (1) workspace por job + descarte controlado (2) apagar stage1/2 após sucesso
# (3) retenção por estado: running nunca, finished 72h, failed 24h (4) ZIP final único
# Checklist: cada job tem pasta própria | stage intermediário descartável | retenção por estado
#   | ZIP final único | sem duplicação | job.json com log resumido | limpeza automática ativa
# OUTPUT_PATH para NC: mesmo critério do app (ARTESP_OUTPUT_DIR ou defaults).
NC_SUBDIRS = ("input", "stage1", "stage2", "final")


def _nc_output_path() -> Path:
    """Base de outputs para NC — mesmo critério do app (OUTPUT_PATH)."""
    env = (os.getenv("ARTESP_OUTPUT_DIR") or "").strip()
    if env:
        return Path(env).resolve()
    if platform.system() == "Linux":
        return Path("/data/outputs").resolve()
    return (Path(__file__).resolve().parent.parent / "outputs").resolve()


def _safe_nc_job_dir(job_id: str) -> Path:
    """
    Retorna o diretório do job NC garantindo que está sob OUTPUT_PATH/nc/.
    Valida path traversal: job_id não pode conter .., / ou \\. Usa .resolve() e
    relative_to(base) para garantir que job_dir fica dentro de OUTPUT_PATH/nc.
    """
    if not job_id or ".." in job_id or "/" in job_id or "\\" in job_id:
        raise HTTPException(status_code=400, detail="job_id inválido.")
    base = _nc_output_path() / "nc"
    job_dir = (base / job_id).resolve()
    base_resolved = base.resolve()
    try:
        job_dir.relative_to(base_resolved)
    except ValueError:
        raise HTTPException(status_code=400, detail="job_id inválido.")
    return job_dir


@dataclass(frozen=True)
class NCWorkspace:
    """Workspace por execução (job): job_dir + subpastas input/, stage1/, stage2/, final/."""

    job_id: str
    job_dir: Path
    input: Path
    stage1: Path
    stage2: Path
    final: Path

    def ensure_dirs(self) -> None:
        for p in (self.input, self.stage1, self.stage2, self.final):
            p.mkdir(parents=True, exist_ok=True)


def _gerar_job_id() -> str:
    """Identificador único por execução: nc_YYYYMMDD_HHMMSS_<suffix> (auditoria/rastreabilidade)."""
    ts = datetime.now(timezone.utc).strftime("%Y%m%d_%H%M%S")
    suffix = secrets.token_hex(4)
    return f"nc_{ts}_{suffix}"


def create_nc_workspace() -> NCWorkspace:
    """
    Cria um workspace por execução (job_id) sob OUTPUT_PATH/nc/<job_id>/.
    Subpastas: input/, stage1/, stage2/, final/.
    Pipeline stateful: arquivos podem ser guardados entre etapas sem re-upload.
    """
    job_id = _gerar_job_id()
    job_dir = _safe_nc_job_dir(job_id)
    job_dir.mkdir(parents=True, exist_ok=True)
    ws = NCWorkspace(
        job_id=job_id,
        job_dir=job_dir,
        input=job_dir / "input",
        stage1=job_dir / "stage1",
        stage2=job_dir / "stage2",
        final=job_dir / "final",
    )
    ws.ensure_dirs()
    _update_job_json(ws, status="created")
    logger.info("NC workspace criado: job_id=%s path=%s", job_id, job_dir)
    return ws


def resolve_nc_workspace(job_id: str) -> NCWorkspace:
    """Resolve workspace existente por job_id (não cria diretórios)."""
    job_dir = _safe_nc_job_dir(job_id)
    return NCWorkspace(
        job_id=job_id,
        job_dir=job_dir,
        input=job_dir / "input",
        stage1=job_dir / "stage1",
        stage2=job_dir / "stage2",
        final=job_dir / "final",
    )


def resolve_workspace(job_id: Optional[str] = None) -> tuple:
    """
    Regra de ouro: etapa isolada cria job / pipeline reutiliza job.
    - job_id None ou "" → cria workspace novo (etapa isolada), retorna (ws, True).
    - job_id preenchido → abre workspace existente (pipeline), touch, retorna (ws, False).
    """
    if job_id and str(job_id).strip():
        j = str(job_id).strip()
        ws = resolve_nc_workspace(j)
        if not ws.job_dir.is_dir():
            raise HTTPException(404, detail="Workspace não encontrado. Execute a primeira etapa (ex.: separar) sem job_id.")
        _update_job_json(ws, status="running")
        return ws, False
    ws = create_nc_workspace()
    _update_job_json(ws, status="running", stage="stage1")
    return ws, True


def _artifacts_for_stage(path: Path, base: Path) -> List[str]:
    """Lista paths relativos (sempre com /) de arquivos em path. Nunca retorna absolutos."""
    if not path.is_dir():
        return []
    out = []
    for p in path.rglob("*"):
        if p.is_file():
            try:
                rel = p.relative_to(base)
                # Normaliza barras (Windows) e evita vazar estrutura do servidor
                out.append(str(rel).replace("\\", "/"))
            except ValueError:
                out.append(p.name)
    return sorted(out)


def _nc_response(
    ws: NCWorkspace,
    stage: str,
    *,
    download_urls: Optional[List[str]] = None,
    final_files: Optional[List[str]] = None,
    artifacts: Optional[Dict[str, List[str]]] = None,
    step_label: Optional[str] = None,
    next_step_label: Optional[str] = None,
) -> Dict[str, Any]:
    """Resposta padrão da API NC: job_id, stage, artifacts, download_urls. step_label/next_step_label = nomes das etapas (iguais aos botões) para exibir no frontend."""
    prefix = f"/outputs/nc/{ws.job_id}"
    if artifacts is None:
        artifacts = {}
    for name, dir_path in [("stage1", ws.stage1), ("stage2", ws.stage2), ("final", ws.final)]:
        if name not in artifacts and dir_path.is_dir():
            artifacts[name] = _artifacts_for_stage(dir_path, ws.job_dir)
    payload: Dict[str, Any] = {
        "ok": True,
        "job_id": ws.job_id,
        "stage": stage,
        "artifacts": artifacts,
    }
    if download_urls:
        payload["download_urls"] = [u if u.startswith("/") else f"{prefix}/{u}".lstrip("/") for u in download_urls]
    if final_files:
        payload["final_files"] = final_files
    if step_label is not None:
        payload["step_label"] = step_label
    if next_step_label is not None:
        payload["next_step_label"] = next_step_label
    return payload


def _list_stage_files(path: Path) -> List[str]:
    """Lista nomes de arquivos em um stage (para job.json)."""
    if not path.is_dir():
        return []
    return sorted(p.name for p in path.iterdir() if p.is_file())


def _update_job_json(
    ws: NCWorkspace,
    status: Optional[str] = None,
    stage: Optional[str] = None,
    stages: Optional[Dict[str, List[str]]] = None,
    log_summary: Optional[Dict[str, Any]] = None,
    retain_hours: Optional[float] = None,
) -> None:
    """
    Atualiza job.json: last_access (limpeza), status, stage, stages, log resumido, retain_until.
    Regra: nunca depender de memória; estado explícito no disco.
    """
    job_json = ws.job_dir / "job.json"
    now = time.time()
    now_iso = datetime.now(timezone.utc).isoformat()
    data: Dict[str, Any] = {
        "job_id": ws.job_id,
        "last_access": now,
        "last_access_iso": now_iso,
    }
    if job_json.is_file():
        try:
            with open(job_json, "r", encoding="utf-8") as f:
                data = json.load(f)
        except (json.JSONDecodeError, OSError):
            pass
    data["last_access"] = now
    data["last_access_iso"] = now_iso
    if "created_at" not in data:
        data["created_at"] = now
    if "created_at_iso" not in data:
        data["created_at_iso"] = datetime.now(timezone.utc).isoformat()
    if status is not None:
        data["status"] = status
    if stage is not None:
        data["stage"] = stage
    if log_summary is not None:
        data["log"] = log_summary
    if retain_hours is not None and retain_hours > 0:
        retain_until = datetime.now(timezone.utc) + timedelta(hours=retain_hours)
        data["retain_until"] = retain_until.isoformat()
        data["retain_until_ts"] = (now + retain_hours * 3600)
    if stages is not None:
        data["stages"] = stages
    else:
        data["stages"] = {
            "input": _list_stage_files(ws.input),
            "stage1": _list_stage_files(ws.stage1),
            "stage2": _list_stage_files(ws.stage2),
            "final": _list_stage_files(ws.final),
        }
    try:
        with open(job_json, "w", encoding="utf-8") as f:
            json.dump(data, f, indent=2)
    except OSError as e:
        logger.warning("Não foi possível atualizar job.json: %s", e)


def _touch_job_access(ws: NCWorkspace) -> None:
    """Atualiza last_access em job.json (compatibilidade; usa _update_job_json)."""
    _update_job_json(ws)


def _safe_input_filename(nome: str) -> str:
    """Nome seguro para salvar em input/ (evita path traversal e caracteres inválidos)."""
    if not nome or ".." in nome or "/" in nome or "\\" in nome:
        return "arquivo.xlsx"
    # mantém só caracteres seguros
    safe = "".join(c for c in nome if c.isalnum() or c in "._- ")
    return safe.strip() or "arquivo.xlsx"


def _purge_dir_contents(path: Path) -> None:
    """Remove todo o conteúdo de um diretório (mantém a pasta). Facilita descarte de stage1/ e stage2/."""
    if not path.is_dir():
        return
    for p in path.iterdir():
        try:
            if p.is_file():
                p.unlink()
            elif p.is_dir():
                shutil.rmtree(p)
        except OSError as e:
            logger.warning("Purge %s: %s", p, e)


def _purge_work_if_finished(ws: NCWorkspace) -> None:
    """Remove stage2/_work para economizar disco quando o job foi finalizado."""
    work = ws.stage2 / "_work"
    if work.is_dir():
        try:
            shutil.rmtree(work)
        except OSError as e:
            logger.warning("Purge _work %s: %s", work, e)


def _garantir_path_nc() -> None:
    """
    Garante que o repositório e nc_artesp estejam em sys.path para:
    - 'from nc_artesp ...' (precisa do repo root)
    - 'from config import ...' / 'from utils ...' dentro dos módulos (precisa de nc_artesp).
    """
    repo = str(_REPO_ROOT)
    proj = str(_NC_PROJ)
    if repo not in sys.path:
        sys.path.insert(0, repo)
    if proj not in sys.path:
        sys.path.insert(0, proj)

# Nomes dos templates conforme config.py e arquivos físicos em nc_artesp/assets/
_NOME_MODELO_KRIA = "Modelo Abertura Evento Kria Conserva Rotina.xlsx"
_NOME_MODELO_RESP = "Modelo.xlsx"
_NOME_MODELO_KCOR = "_Planilha Modelo Kcor-Kria.XLSX"


def _ler_asset(nome: str, pasta_base: Path | None = None) -> bytes:
    """Lê template de nc_artesp ou de pasta_base (ex.: Artemig). pasta_base=None → só nc_artesp."""
    pastas = ([pasta_base] if pasta_base and pasta_base.is_dir() else []) + [_NC_ASSETS, _NC_ASSETS.parent]
    for pasta in pastas:
        if not pasta or not pasta.is_dir():
            continue
        p = pasta / nome
        if p.is_file():
            return p.read_bytes()
        nome_lower = nome.lower()
        for f in pasta.iterdir():
            if f.is_file() and f.name.lower() == nome_lower:
                return f.read_bytes()
    raise HTTPException(
        status_code=503,
        detail=f"Template '{nome}' não encontrado em nc_artesp/assets/templates/ ou assets/.",
    )


def _carregar_modelo_kria(lote: str | None = None) -> bytes:
    if (lote or "").strip() == "50":
        try:
            return _ler_asset(_NOME_MODELO_KRIA, _NC_ARTEMIG_TEMPLATES)
        except HTTPException:
            pass
    return _ler_asset(_NOME_MODELO_KRIA)


def _carregar_modelo_resp(lote: str | None = None) -> bytes:
    pasta = _NC_ARTEMIG_TEMPLATES if (lote or "").strip() == "50" else None
    for nome in (_NOME_MODELO_RESP, "Modelo Resposta.xlsx", "Modelo_Resposta.xlsx"):
        try:
            return _ler_asset(nome, pasta) if pasta and pasta.is_dir() else _ler_asset(nome)
        except HTTPException:
            continue
    raise HTTPException(
        status_code=503,
        detail=f"Template de Resposta não encontrado (tentou {_NOME_MODELO_RESP} e alternativas).",
    )


def _carregar_modelo_kcor(lote: str | None = None) -> bytes:
    if (lote or "").strip() == "50":
        try:
            from nc_artemig.config import MODELO_KCOR_KRIA

            p = Path(MODELO_KCOR_KRIA)
            if p.is_file():
                return p.read_bytes()
        except Exception:
            pass
        try:
            return _ler_asset(_NOME_MODELO_KCOR, _NC_ARTEMIG_TEMPLATES)
        except HTTPException:
            pass
    return _ler_asset(_NOME_MODELO_KCOR)


def _check_auth(request: Request) -> dict:
    """Auth desabilitado para NC — acesso direto sem login."""
    return {}


def _ler(f: UploadFile) -> bytes:
    data = f.file.read()
    if len(data) > MAX_BYTES:
        raise HTTPException(413, f"Arquivo '{f.filename}' excede {MAX_MB} MB.")
    return data


def _safe_filename_header(nome: str) -> str:
    """Nome seguro para Content-Disposition (evita erro latin-1 em headers HTTP)."""
    if not nome:
        return nome
    nfd = unicodedata.normalize("NFD", nome)
    sem_comb = "".join(c for c in nfd if unicodedata.category(c) != "Mn")
    try:
        return sem_comb.encode("latin-1").decode("latin-1")
    except UnicodeEncodeError:
        return sem_comb.encode("latin-1", "replace").decode("latin-1")


def _stream_zip(data: bytes, nome: str, job_id: Optional[str] = None) -> StreamingResponse:
    """Stream ZIP; se job_id informado, adiciona header X-NC-Job-Id para o front encadear."""
    headers = {"Content-Disposition": f'attachment; filename="{_safe_filename_header(nome)}"'}
    if job_id:
        headers["X-NC-Job-Id"] = job_id
    return StreamingResponse(
        io.BytesIO(data), media_type="application/zip",
        headers=headers,
    )


def _stream_xlsx(data: bytes, nome: str, job_id: Optional[str] = None) -> StreamingResponse:
    """Stream XLSX; se job_id informado, adiciona header X-NC-Job-Id para o front encadear."""
    headers = {"Content-Disposition": f'attachment; filename="{_safe_filename_header(nome)}"'}
    if job_id:
        headers["X-NC-Job-Id"] = job_id
    return StreamingResponse(
        io.BytesIO(data),
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers=headers,
    )


def _nc_proj_disponivel() -> bool:
    return _NC_PROJ.exists()


_modulos_carregados_log: set = set()


def _importar_modulo(nome: str):
    """
    Importa módulo do nc_artesp adicionando ao sys.path se necessário.
    Funciona tanto no Render (nc_artesp/ no repo) quanto no desktop Windows.
    """
    _garantir_path_nc()
    if not _nc_proj_disponivel():
        raise HTTPException(
            503,
            f"Módulos NC não encontrados.\n"
            f"Caminho verificado: {_NC_PROJ}\n"
            f"Opções:\n"
            f"  1. Copie artesp_nc_v2.0/modulos/ para GeradorARTESP/nc_artesp/modulos/\n"
            f"  2. Defina a variável de ambiente ARTESP_NC_PROJ no Render.",
        )
    proj = str(_NC_PROJ)
    if proj not in sys.path:
        sys.path.insert(0, proj)
    try:
        import importlib
        mod = importlib.import_module(f"modulos.{nome}")
        if nome not in _modulos_carregados_log:
            _modulos_carregados_log.add(nome)
            mod_path = getattr(mod, "__file__", "?")
            logger.info("NC módulo carregado: %s → %s | pasta NC: %s", nome, mod_path, _NC_PROJ)
        return mod
    except ImportError as e:
        raise HTTPException(503, f"Módulo '{nome}' não carregado: {e}")


router = APIRouter(prefix="/nc", tags=["NC Artesp"])


def _importar_analisar_pdf():
    """Importa módulo de análise de PDF de NC."""
    _garantir_path_nc()
    try:
        from nc_artesp.modulos import analisar_pdf_nc
        return analisar_pdf_nc
    except ImportError as e:
        raise HTTPException(
            status_code=503,
            detail=(
                f"Módulo de análise não disponível: {e}\n"
                "Verifique se pymupdf e reportlab estão instalados."
            ),
        )


def _importar_pdf_extractor():
    """Importa pdf_extractor de nc_artesp (raiz do repositório)."""
    _garantir_path_nc()
    try:
        from nc_artesp import pdf_extractor
        return pdf_extractor
    except ImportError as e:
        raise HTTPException(
            status_code=503,
            detail=(
                f"Extração de PDF não disponível: {e}\n"
                "Verifique se pymupdf e pillow estão instalados."
            ),
        )


@router.post(
    "/analisar-pdf",
    summary="Analisar sequência de KMs e tipos de NCs do PDF de Constatação",
    response_description="ZIP com PDF de análise e XLSX no formato do template de Relatório de Fiscalização",
)
async def nc_analisar_pdf(
    request: Request,
    pdfs: List[UploadFile] = File(..., description="Um ou mais PDFs de NC Constatação de Rotina Artesp"),
    limiar_km: float = Form(2.0, description="Gap mínimo de KM para gerar alerta (padrão 2 km)"),
    lote: str = Form("", description="Lote para o relatório (13, 21, 26 ou 50 Artemig). Obrigatório."),
    excel: List[UploadFile] = File(default=[], description="Um ou mais Excels que acompanham os PDFs (mesmo layout do relatório). Preenchem col E, O, P."),
):
    """
    Analisa **um ou mais** PDFs de Constatação de Rotina Artesp.
    Se **excel** for enviado (um ou mais arquivos no mesmo formato do relatório), usa as colunas E, O e P para preencher/complementar os dados extraídos dos PDFs.
    Retorna **ZIP** com PDF de análise e XLSX do relatório.
    """
    _check_auth(request)
    if not (lote or "").strip():
        raise HTTPException(400, "Selecione o lote.")
    mod = _importar_analisar_pdf()
    try:
        pdfs_bytes = []
        for i, f in enumerate(pdfs):
            data = await f.read()
            if len(data) > MAX_BYTES:
                raise HTTPException(413, f"Arquivo '{f.filename}' excede {MAX_MB} MB.")
            pdfs_bytes.append(data)
        nomes = [Path(f.filename or f"pdf_{i+1}").stem for i, f in enumerate(pdfs)]
        lote_ok = (lote or "").strip() or None
        excel_list: List[bytes] = []
        for f in excel or []:
            if f and f.filename and (f.filename.lower().endswith(".xlsx") or f.filename.lower().endswith(".xls")):
                data = await f.read()
                if len(data) > MAX_BYTES:
                    raise HTTPException(413, f"Arquivo Excel '{f.filename}' excede o tamanho máximo.")
                excel_list.append(data)
        pdf_rel, xlsx_bytes, resumo = await asyncio.to_thread(
            mod.analisar_e_gerar_pdf_multi, pdfs_bytes, limiar_km, nomes, lote_ok, excel_list
        )
        n_arqs = len(pdfs)
        relatorio_hoje = resumo.get("relatorio_hoje", True)
        if relatorio_hoje:
            n_emerg = len(resumo.get("emergenciais_lista", []))
            n_alertas = len(resumo.get("alertas_gap", []))
            n_ocultos = resumo.get("total_ocultos", 0)
        else:
            n_emerg = n_alertas = n_ocultos = 0
            logger.info(
                "nc/analisar-pdf: relatório anterior (data constatação ≠ hoje); "
                "alertas não exibidos no PDF (Total NCs: %s)",
                resumo.get("total", 0),
            )
        # Nome do ZIP alinhado ao lote do formulário (evita slug 13 quando o cliente envia 50).
        lote_slug = (lote or "").strip() or "13"
        try:
            _rotulo, slug = mod.rotulo_e_slug_lote_para_saida(lote_slug)
        except Exception:
            slug = (resumo.get("slug_zip") or "Lote13_Rodovias_Colinas").strip()
        slug = "".join(c if c.isalnum() or c in "_-" else "_" for c in slug) or "Analise"
        pasta = slug
        nome_pdf = f"{pasta}/Analise_NCs_{slug}.pdf"
        nome_xlsx = f"{pasta}/Relatorio_Fiscalizacao_{slug}.xlsx"
        nome_zip = f"Relatorio_Analise_NCs_{slug}.zip"
        buf = io.BytesIO()
        with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
            zf.writestr(nome_pdf.replace("\\", "/"), pdf_rel)
            zf.writestr(nome_xlsx.replace("\\", "/"), xlsx_bytes)
            kcor_b = resumo.get("exportar_kcor_xlsx") or b""
            kcor_nome = (resumo.get("exportar_kcor_nome") or "").strip()
            if kcor_b and kcor_nome and (lote_slug or "").strip() == "50":
                arc_kcor = f"{pasta}/{kcor_nome}".replace("\\", "/")
                zf.writestr(arc_kcor, kcor_b)
        zip_bytes = buf.getvalue()
        return Response(
            content=zip_bytes,
            media_type="application/zip",
            headers={
                "Content-Disposition": f'attachment; filename="{nome_zip}"',
                "X-NC-Total":           str(resumo.get("total", 0)),
                "X-NC-Emergenciais":    str(n_emerg),
                "X-NC-Alertas":         str(n_alertas),
                "X-NC-Ocultos":         str(n_ocultos),
                "X-NC-Arquivos":        str(n_arqs),
                "X-NC-Relatorio-Dia":   "1" if relatorio_hoje else "0",
                "X-NC-Zip-Slug":        slug,
            },
        )
    except HTTPException:
        raise
    except ValueError as e:
        raise HTTPException(400, str(e))
    except Exception as e:
        logger.error("nc/analisar-pdf: %s", traceback.format_exc())
        raise HTTPException(500, str(e))


@router.post(
    "/extrair-pdf",
    summary="Extrair imagens do PDF de NC Constatação",
    response_description="ZIP: lote 50 = arquivos na raiz; outros lotes = nc/ e PDF/",
)
async def nc_extrair_pdf(
    request: Request,
    pdfs: List[UploadFile] = File(..., description="Um ou mais PDFs de NC Constatação Artesp"),
    dpi: Optional[int] = Form(
        None,
        description="Resolução PyMuPDF antes do redimensionamento; padrão = ARTESP_M02_EXTRACAO_RENDER_DPI (150).",
    ),
    lote: str = Form("", description="Número do lote (13, 21, 26, 50…). Obrigatório; entra no nome do ZIP."),
    nomear_por_indice_fiscalizacao: bool = Form(
        False,
        description="Se True, nomeia fotos por índice (00001, 00002...) em vez do código do PDF. Use para Meio Ambiente / Modelo Foto Kria.",
    ),
):
    """
    Processa **um ou mais** PDFs de Não Conformidade Artesp e gera dois arquivos JPG por NC:
    - **nc(N).jpg** — apenas a foto
    - **PDF(N).jpg** — texto de cabeçalho + foto

    Se nomear_por_indice_fiscalizacao=True, N = 00001, 00002... (índice = coluna V da EAF).
    Caso contrário, N = código extraído do PDF (ex.: 896643, HE.13.0111).
    **Lote 50 (Artemig):** um único nível no ZIP — **Texto (COD).jpg**, **PDF (COD).jpg**, **nc (COD).jpg** na raiz.
    Demais lotes: pastas **nc/** e **PDF/** como antes.
    """
    _check_auth(request)
    if not (lote or "").strip():
        raise HTTPException(400, "Selecione o lote.")
    extrator = _importar_pdf_extractor()
    lote_m = re.search(r"\d+", (lote or "").strip() or "13")
    pasta_unica_artemig = (lote_m.group(0) if lote_m else "") == "50"
    try:
        arquivos: dict[str, bytes] = {}

        for f in pdfs:
            pdf_bytes = _ler(f)
            zip_bytes_i, _ = extrator.extrair_pdf_para_zip(
                pdf_bytes,
                dpi=dpi,
                nomear_por_indice_fiscalizacao=nomear_por_indice_fiscalizacao,
                pasta_unica=pasta_unica_artemig,
            )
            with zipfile.ZipFile(io.BytesIO(zip_bytes_i)) as zf_i:
                for name in zf_i.namelist():
                    if name.endswith("/") or not name.strip():
                        continue
                    final = name.replace("\\", "/")
                    n_col = 1
                    while final in arquivos:
                        pth = Path(name)
                        stem, suf = pth.stem, (pth.suffix or ".jpg")
                        par = pth.parent
                        stem2 = f"{stem}_{n_col}"
                        final = (
                            f"{par.as_posix()}/{stem2}{suf}"
                            if str(par) != "."
                            else f"{stem2}{suf}"
                        )
                        n_col += 1
                    arquivos[final] = zf_i.read(name)

        buf = io.BytesIO()
        with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf_out:
            for nome, data in arquivos.items():
                zf_out.writestr(nome, data)

        n = len(pdfs)
        lote_num = lote_m.group(0) if lote_m else "13"
        nome_zip = f"lote_{lote_num}_{n}_pdfs_imagens.zip"
        return _stream_zip(buf.getvalue(), nome_zip)
    except HTTPException:
        raise
    except Exception as e:
        logger.error("nc/extrair-pdf: %s", traceback.format_exc())
        raise HTTPException(500, str(e))


@router.post("/separar", summary="M01 — Separar NC da planilha EAF")
async def nc_separar(
    request: Request,
    eafs: List[UploadFile] = File(..., description="Uma ou mais planilhas EAF (.xlsx ou .xls)"),
    job_id: Optional[str] = Form(None),
    finalize: bool = Form(False),
):
    """
    Etapa isolada: sem job_id → cria job, grava stage1/, marca finished, retorna job_id e link.
    Pipeline: com job_id → reutiliza job, grava em stage1/, não marca finished (a menos que finalize=1).
    """
    _check_auth(request)
    mod = _importar_modulo("separar_nc")
    try:
        ws, created = resolve_workspace(job_id or None)
        stage1_dir = ws.stage1
        stage1_dir.mkdir(parents=True, exist_ok=True)
        todos: list[tuple[str, bytes]] = []   # (nome_arquivo, bytes)
        for idx, eaf in enumerate(eafs):
            eaf_bytes = _ler(eaf)
            nome_eaf = eaf.filename or "eaf.xlsx"
            eaf_path = ws.input / nome_eaf
            ws.input.mkdir(parents=True, exist_ok=True)
            eaf_path.write_bytes(eaf_bytes)
            pasta_dest = stage1_dir / f"{DIR_EXPORTAR_EAF}_{idx}"
            pasta_dest.mkdir(parents=True, exist_ok=True)
            arqs = mod.executar(eaf_path, pasta_destino=pasta_dest)
            for a in (arqs or []):
                p = Path(a)
                if p.is_file():
                    todos.append((p.name, p.read_bytes()))

        buf  = io.BytesIO()
        seen: set[str] = set()
        with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
            for nome, data in todos:
                final = nome
                n = 1
                while final in seen:
                    stem = Path(nome).stem
                    ext  = Path(nome).suffix
                    final = f"{stem}_{n}{ext}"
                    n += 1
                seen.add(final)
                zf.writestr(final, data)

        zip_bytes = buf.getvalue()
        ws.final.mkdir(parents=True, exist_ok=True)
        (ws.final / "nc_separados.zip").write_bytes(zip_bytes)

        is_final = created or finalize
        if is_final:
            retain_hours = 24 if created else 72  # etapa isolada 24h, pipeline 72h
            _update_job_json(ws, status="finished", stage="final", retain_hours=retain_hours)
        else:
            _update_job_json(ws, status="running", stage="stage1")

        if request.query_params.get("format") == "json":
            return JSONResponse(_nc_response(
                ws, "final" if is_final else "stage1",
                download_urls=["final/nc_separados.zip"],
                final_files=["nc_separados.zip"] if is_final else None,
                step_label="Separar NC",
                next_step_label="E-mail, Modelo Foto",
            ))
        return _stream_zip(zip_bytes, "nc_separados.zip", ws.job_id)
    except HTTPException:
        raise
    except Exception as e:
        logger.error("nc/separar: %s", traceback.format_exc())
        raise HTTPException(500, str(e))


@router.post("/criar-email", summary="Gerar e-mails .eml a partir do ZIP de XLS (saída Separar NC)")
async def nc_criar_email_endpoint(
    request: Request,
    xls_zip: UploadFile = File(..., description="ZIP com XLS individuais (saída de Separar NC)"),
    imagens_pdf_zip: Optional[UploadFile] = File(None, description="ZIP opcional com imagens PDF (N).jpg (saída Extrair PDF) — para embutir fotos nos e-mails"),
):
    """
    Gera arquivos .eml (rascunhos de resposta NC) a partir do ZIP de planilhas XLS.
    Se enviar imagens_pdf_zip, as fotos são embutidas no corpo do e-mail.
    Retorna um ZIP com todos os .eml gerados (pasta emails/).
    """
    _check_auth(request)
    _garantir_path_nc()
    mod = _importar_modulo("nc_criar_email")
    try:
        xls_bytes = await xls_zip.read()
        imagens_bytes = await imagens_pdf_zip.read() if (imagens_pdf_zip and imagens_pdf_zip.filename) else None
        if len(xls_bytes) > MAX_BYTES:
            raise HTTPException(413, f"Arquivo '{xls_zip.filename}' excede {MAX_MB} MB.")
        if imagens_bytes is not None and len(imagens_bytes) > MAX_BYTES:
            raise HTTPException(413, f"Arquivo '{imagens_pdf_zip.filename}' excede {MAX_MB} MB.")
        with tempfile.TemporaryDirectory(prefix="nc_email_") as tmp:
            tmp_path = Path(tmp)
            pasta_xls = tmp_path / DIR_EXPORTAR
            pasta_xls.mkdir(parents=True, exist_ok=True)
            with zipfile.ZipFile(io.BytesIO(xls_bytes), "r") as zf:
                zf.extractall(str(pasta_xls))

            pasta_fotos_pdf = tmp_path / DIR_IMAGENS_PDF
            if imagens_bytes is not None:
                pasta_fotos_pdf.mkdir(parents=True, exist_ok=True)
                with zipfile.ZipFile(io.BytesIO(imagens_bytes), "r") as zf:
                    zf.extractall(str(pasta_fotos_pdf))

            pasta_emails = tmp_path / "emails"
            pasta_emails.mkdir(parents=True, exist_ok=True)
            resultado = await asyncio.to_thread(
                mod.executar,
                pasta_xls=pasta_xls,
                pasta_fotos_pdf=pasta_fotos_pdf if pasta_fotos_pdf.is_dir() else None,
                usar_outlook=False,
                pasta_saida_eml=pasta_emails,
            )
            emls = resultado.get("eml") or []
            buf = io.BytesIO()
            with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
                for eml in pasta_emails.rglob("*"):
                    if eml.is_file():
                        zf.write(eml, f"emails/{eml.name}")
            zip_bytes = buf.getvalue()
        if len(emls) == 0:
            logger.warning("criar-email: nenhum .eml gerado. Verifique o ZIP (XLS com colunas C, U, V preenchidas).")
        return _stream_zip(zip_bytes, "emails_nc.zip")
    except HTTPException:
        raise
    except Exception as e:
        logger.error("nc/criar-email: %s", traceback.format_exc())
        raise HTTPException(500, str(e))


@router.post("/gerar-modelo-foto", summary="M02 — Gerar Kria + Resposta")
async def nc_gerar_modelo_foto(
    request: Request,
    xls_zip: UploadFile = File(..., description="ZIP com os XLS individuais (saída M01)"),
    modelo_kria: Optional[UploadFile] = File(None, description="Modelo Kria (.xlsx) — padrão: assets/Modelo Abertura Evento Kria Conserva Rotina.xlsx"),
    modelo_resp: Optional[UploadFile] = File(None, description="Modelo Resposta (.xlsx) — padrão: assets/Modelo.xlsx"),
    fotos_pdf_zip: Optional[UploadFile] = File(None, description="ZIP com fotos PDF (N).jpg (opcional — saída do Extrair PDF)"),
    job_id: Optional[str] = Form(None),
    finalize: bool = Form(False),
    lote: Optional[str] = Form(None, description="Lote 50 = ARTEMIG (templates em nc_artemig/assets/Template)"),
):
    """
    Etapa isolada: sem job_id → cria job, grava stage2/ e final/, marca finished.
    Pipeline: com job_id → reutiliza job, grava em stage2/; finalize=1 marca finished.
    """
    _check_auth(request)
    mod = _importar_modulo("gerar_modelo_foto")
    lote_ok = (lote or "").strip() or None
    try:
        ws, created = resolve_workspace(job_id or None)
        work = ws.stage2 / "_work"
        work.mkdir(parents=True, exist_ok=True)

        pasta_xls = work / DIR_EXPORTAR
        pasta_xls.mkdir(parents=True, exist_ok=True)
        _purge_dir_contents(pasta_xls)
        with zipfile.ZipFile(io.BytesIO(_ler(xls_zip))) as zf:
            zf.extractall(str(pasta_xls))

        p_modelo_kria = work / "modelo_kria.xlsx"
        p_modelo_resp = work / "modelo_resp.xlsx"
        p_modelo_kria.write_bytes(_ler(modelo_kria) if modelo_kria else _carregar_modelo_kria(lote_ok))
        p_modelo_resp.write_bytes(_ler(modelo_resp) if modelo_resp else _carregar_modelo_resp(lote_ok))

        pasta_fotos_pdf = None
        pasta_fotos_nc = None
        if fotos_pdf_zip:
            pasta_fotos_pdf = work / DIR_IMAGENS_PDF
            pasta_fotos_pdf.mkdir(parents=True, exist_ok=True)
            _purge_dir_contents(pasta_fotos_pdf)
            with zipfile.ZipFile(io.BytesIO(_ler(fotos_pdf_zip))) as zf:
                zf.extractall(str(pasta_fotos_pdf))
            # O ZIP do Extrair PDF contém nc (CODIGO).jpg e PDF (CODIGO).jpg na mesma pasta
            pasta_fotos_nc = pasta_fotos_pdf

        # Resolver como absolutos; limpar saídas anteriores para devolver só os arquivos desta requisição
        pasta_kria = (work / DIR_KRIA).resolve()
        pasta_resp = (work / DIR_RESPOSTAS_PENDENTES).resolve()
        pasta_kria.mkdir(parents=True, exist_ok=True)
        pasta_resp.mkdir(parents=True, exist_ok=True)
        _purge_dir_contents(pasta_kria)
        _purge_dir_contents(pasta_resp)

        resultado = mod.executar(
            pasta_xls=pasta_xls.resolve() if hasattr(pasta_xls, "resolve") else pasta_xls,
            modelo_kria=p_modelo_kria,
            pasta_saida_kria=pasta_kria,
            modelo_resposta=p_modelo_resp,
            pasta_saida_resp=pasta_resp,
            pasta_fotos_nc=pasta_fotos_nc,
            pasta_fotos_pdf=pasta_fotos_pdf,
        )

        buf = io.BytesIO()
        n_arquivos_zip = 0
        with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
            for d in [pasta_kria, pasta_resp]:
                for f in sorted(d.rglob("*.xlsx")):
                    if f.is_file():
                        try:
                            zf.write(f, f"{d.name}/{f.name}")
                            n_arquivos_zip += 1
                        except Exception as ex:
                            logger.warning("gerar-modelo-foto: não foi possível adicionar ao ZIP %s: %s", f, ex)
        zip_bytes = buf.getvalue()

        # Nunca devolver ZIP vazio: rejeitar por contagem ou por tamanho (ZIP vazio ≈ 22 bytes)
        if n_arquivos_zip == 0 or len(zip_bytes) < 100:
            entradas = list(pasta_xls.rglob("*.xls*"))
            entradas = [f for f in entradas if f.is_file()]
            n_erros = len(resultado.get("erros") or [])
            msg = (
                "Nenhum arquivo gerado (ZIP vazio). "
                "Verifique: 1) O ZIP deve conter planilhas .xlsx ou .xls (saída do M01 Separar NC ou EAF); "
                "2) Planilhas com coluna C (Código) preenchida a partir da linha 5."
            )
            if not entradas:
                msg += f" Nenhuma planilha .xls/.xlsx encontrada na pasta extraída ({pasta_xls})."
            else:
                msg += f" Encontradas {len(entradas)} planilha(s) de entrada; nenhuma NC válida ou erro ao processar."
                if n_erros:
                    msg += f" Erros em {n_erros} arquivo(s): {resultado.get('erros', [])[:3]}."
            msg += " Confira os logs do servidor para detalhes (ex.: Permission denied, modelo não encontrado)."
            logger.warning("gerar-modelo-foto: %s", msg)
            raise HTTPException(status_code=422, detail=msg)
        ws.final.mkdir(parents=True, exist_ok=True)
        (ws.final / "modelos_kria.zip").write_bytes(zip_bytes)

        is_final = created or finalize
        if is_final:
            retain_hours = 24 if created else 72  # etapa isolada 24h, pipeline 72h
            _update_job_json(ws, status="finished", stage="final", retain_hours=retain_hours)
            _purge_work_if_finished(ws)
        else:
            _update_job_json(ws, status="running", stage="stage2")

        if request.query_params.get("format") == "zip":
            return _stream_zip(zip_bytes, "modelos_kria.zip", ws.job_id)
        return JSONResponse(_nc_response(
            ws, "final" if is_final else "stage2",
            download_urls=["final/modelos_kria.zip"],
            final_files=["modelos_kria.zip"] if is_final else None,
            step_label="Modelo Foto",
            next_step_label="Conservação",
        ))
    except HTTPException:
        raise
    except Exception as e:
        logger.error("nc/gerar-modelo-foto: %s", traceback.format_exc())
        raise HTTPException(500, str(e))


@router.post("/inserir-conservacao", summary="M03 — Kcor-Kria Conservação")
async def nc_inserir_conservacao(
    request: Request,
    kria_zip: UploadFile = File(..., description="ZIP com planilhas Kria (saída M02)"),
    modelo_kcor: Optional[UploadFile] = File(None, description="Modelo Kcor-Kria (.xlsx) — padrão: assets/_Planilha Modelo Kcor-Kria.XLSX"),
    fotos_zip: Optional[UploadFile] = File(None, description="ZIP fotos PDF (opcional)"),
    job_id: Optional[str] = Form(None),
    finalize: bool = Form(False),
    lote: Optional[str] = Form(None, description="Lote 50 = ARTEMIG (template em nc_artemig/assets/Template)"),
):
    _check_auth(request)
    return await _inserir_nc(request, kria_zip, modelo_kcor, fotos_zip, modo="conservacao", job_id=job_id, finalize=finalize, lote=lote)


@router.post("/inserir-meio-ambiente", summary="M07 — Kcor-Kria Meio Ambiente")
async def nc_inserir_ma(
    request: Request,
    kria_zip: UploadFile = File(..., description="ZIP com planilhas Kria MA"),
    modelo_kcor: Optional[UploadFile] = File(None, description="Modelo Kcor-Kria (.xlsx) — padrão: assets/_Planilha Modelo Kcor-Kria.XLSX"),
    job_id: Optional[str] = Form(None),
    finalize: bool = Form(False),
    lote: Optional[str] = Form(None, description="Lote 50 = ARTEMIG (template em nc_artemig/assets/Template)"),
):
    _check_auth(request)
    return await _inserir_nc(request, kria_zip, modelo_kcor, None, modo="meio_ambiente", job_id=job_id, finalize=finalize, lote=lote)


@router.post("/inserir-meio-ambiente-pdf", summary="M07 — Separar NC Meio Ambiente a partir de PDF(s)")
async def nc_inserir_ma_pdf(
    request: Request,
    pdf: list[UploadFile] = File(..., description="Um ou mais PDFs de Meio Ambiente"),
    job_id: Optional[str] = Form(None),
    finalize: bool = Form(False),
):
    """Processa um ou mais PDFs de Meio Ambiente: extrai as informações em TEXTO do PDF,
    gera a planilha EAF (passo 1) e executa o Separar NC. Retorna ZIP com EAF + NCs separados."""
    _check_auth(request)
    if not pdf:
        raise HTTPException(400, "Envie pelo menos um PDF.")
    pdf_list = pdf if isinstance(pdf, list) else [pdf]
    mod_kria = _importar_modulo("inserir_nc_kria")
    mod_separar = _importar_modulo("separar_nc")
    try:
        ws, created = resolve_workspace(job_id or None)
        list_pdf_bytes = [await up.read() for up in pdf_list]
        download_urls = []
        ws.input.mkdir(parents=True, exist_ok=True)
        ws.stage1.mkdir(parents=True, exist_ok=True)

        eaf_path = mod_kria.gerar_eaf_desde_pdfs_ma(
            list_pdf_bytes,
            pasta_saida=ws.input,
            nome_arquivo="eaf_ma_desde_pdf.xlsx",
        )
        if not eaf_path or not eaf_path.is_file():
            raise HTTPException(
                422,
                "Não foi possível gerar a planilha EAF a partir dos PDFs. Verifique se o PDF é de Meio Ambiente e contém blocos com data, código ou KM.",
            )
        download_urls.append(f"input/{eaf_path.name}")

        arqs = mod_separar.executar(
            arquivo_mae=eaf_path,
            pasta_destino=ws.stage1,
            um_arquivo_por_nc=True,
        )

        # ZIP de saída MA: nc_separados_ma.zip com pasta "Separar NC MA"
        buf = io.BytesIO()
        nome_zip = "nc_separados_ma.zip"
        pasta_raiz = "Separar NC MA"
        with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
            if arqs:
                for a in arqs:
                    p = Path(a)
                    if p.is_file():
                        try:
                            arcname = f"{pasta_raiz}/{p.relative_to(ws.stage1).as_posix()}"
                        except ValueError:
                            arcname = f"{pasta_raiz}/{p.name}"
                        zf.write(p, arcname)
            else:
                zf.write(eaf_path, f"{pasta_raiz}/{eaf_path.name}")
                zf.writestr(
                    f"{pasta_raiz}/LEIA-ME Separar NC.txt",
                    "O Separar NC não gerou arquivos individuais.\n"
                    "A planilha EAF está neste ZIP — abra e verifique se há dados a partir da linha 5 (colunas C e Q).\n"
                    "Se a EAF estiver correta, o problema pode ser o template Template_EAF.xlsx em nc_artesp/assets.",
                )
        zip_bytes = buf.getvalue()
        ws.final.mkdir(parents=True, exist_ok=True)
        (ws.final / nome_zip).write_bytes(zip_bytes)
        download_urls.append(f"final/{nome_zip}")

        is_final = created or finalize
        if is_final:
            _update_job_json(ws, status="finished", stage="final", retain_hours=24 if created else 72)
            _purge_work_if_finished(ws)
        else:
            _update_job_json(ws, status="running", stage="stage1")
        if request.query_params.get("format") == "zip":
            return _stream_zip(zip_bytes, nome_zip, ws.job_id)
        return JSONResponse(_nc_response(
            ws, "final" if is_final else "stage1",
            download_urls=download_urls,
            final_files=[nome_zip] if is_final else None,
            step_label="Meio Ambiente PDF",
            next_step_label="Acumulado",
        ))
    except HTTPException:
        raise
    except Exception as e:
        logger.error("nc/inserir-meio-ambiente-pdf: %s", traceback.format_exc())
        raise HTTPException(500, str(e))


@router.post("/pipeline-meio-ambiente-pdf", summary="M1+M2+M3 — Pipeline Meio Ambiente a partir de PDF")
async def nc_pipeline_ma_pdf(
    request: Request,
    pdf: UploadFile = File(..., description="PDF de Meio Ambiente"),
    imagens_zip: Optional[UploadFile] = File(None, description="ZIP opcional com imagens extraídas (nc (N).jpg, PDF (N).jpg). Se enviado, é extraído no início e usado no M2."),
    job_id: Optional[str] = Form(None),
    finalize: bool = Form(False),
    lote: Optional[str] = Form(None, description="Lote 50 = ARTEMIG (templates em nc_artemig/assets/Template)"),
):
    """
    Executa o equivalente a M1, M2 e M3 a partir do PDF de Meio Ambiente:
    M1 = extrai e parseia NCs do texto do PDF; gera também a planilha EAF (template do Separar NC).
    M2 = gera Kria (Arquivo Foto - MA) e Resposta (Pendentes).
    M3 = gera Kcor-Kria e imagens.
    Opcional: envie imagens_zip com as fotos extraídas (nc (1).jpg, PDF (1).jpg, etc.) para o M2 preencher os modelos.
    Retorna ZIP com EAF MA, Kria MA, Resposta MA, Kcor-Kria Meio Ambiente e Imagens MA.
    """
    _check_auth(request)
    mod = _importar_modulo("inserir_nc_kria")
    try:
        ws, created = resolve_workspace(job_id or None)
        work = ws.stage2 / "_work"
        work.mkdir(parents=True, exist_ok=True)
        work = work.resolve()
        pasta_imagens = (work / DIR_IMAGENS_MA).resolve()
        pasta_kria = (work / "Arquivo Foto MA").resolve()
        pasta_resp = (work / "Resposta MA").resolve()
        pasta_kcor = (work / DIR_MA).resolve()
        pasta_eaf = (work / "EAF MA").resolve()
        pasta_separar_nc = (work / "Separar NC MA").resolve()
        for p in (pasta_imagens, pasta_kria, pasta_resp, pasta_kcor, pasta_eaf, pasta_separar_nc):
            p.mkdir(parents=True, exist_ok=True)
        # Evitar acúmulo (persistência disco entre requisições/job_id reutilizado):
        for p in (pasta_imagens, pasta_kria, pasta_resp, pasta_kcor, pasta_eaf, pasta_separar_nc):
            _purge_dir_contents(p)
        # Opcional: extrair ZIP de imagens no início (nc (N).jpg, PDF (N).jpg) para o M2 usar
        if imagens_zip and imagens_zip.filename and (imagens_zip.filename.lower().endswith(".zip") or getattr(imagens_zip, "content_type", "") == "application/zip"):
            try:
                zip_bytes = await imagens_zip.read()
                if len(zip_bytes) > 0:
                    with zipfile.ZipFile(io.BytesIO(zip_bytes), "r") as zf:
                        zf.extractall(str(pasta_imagens))
                    logger.info("Pipeline MA: ZIP de imagens extraído em %s", pasta_imagens.name)
            except Exception as e_zip:
                logger.warning("Pipeline MA: falha ao extrair ZIP de imagens: %s", e_zip)
        pdf_bytes = await pdf.read()
        nome_origem = (pdf.filename or "PDF MA").replace(".pdf", "").replace(".PDF", "")[:50]
        lote_ok = (lote or "").strip() or None
        p_kria = work / "modelo_kria_ma.xlsx"
        p_resp = work / "modelo_resp_ma.xlsx"
        modelo_kria_ma = None
        modelo_resp_ma = None
        try:
            p_kria.write_bytes(_carregar_modelo_kria(lote_ok))
            modelo_kria_ma = p_kria
        except HTTPException:
            pass
        try:
            p_resp.write_bytes(_carregar_modelo_resp(lote_ok))
            modelo_resp_ma = p_resp
        except HTTPException:
            # Fallback: usar o mesmo modelo do Kria para o Resposta e garantir que o segundo modelo seja gerado
            if modelo_kria_ma and modelo_kria_ma.is_file():
                p_resp.write_bytes(p_kria.read_bytes())
                modelo_resp_ma = p_resp
                logger.warning("Modelo Resposta não encontrado; usando modelo Kria como fallback para gerar o segundo arquivo.")
        # Se ainda faltar modelo Kria (asset não carregou), tentar path do config do nc_artesp
        if modelo_kria_ma is None:
            try:
                _garantir_path_nc()
                import importlib
                cfg = importlib.import_module("config")
                p = getattr(cfg, "M02_MODELO_KRIA", None)
                if p is not None and getattr(p, "is_file", lambda: False)() and getattr(p, "suffix", "") and str(p.suffix).lower() == ".xlsx":
                    modelo_kria_ma = p
                    if not modelo_resp_ma or not getattr(modelo_resp_ma, "is_file", lambda: False)():
                        p_resp.write_bytes(p.read_bytes())
                        modelo_resp_ma = p_resp
                    logger.info("Usando modelo Kria do config nc_artesp para MA.")
            except Exception:
                pass
        if modelo_kria_ma is None or not getattr(modelo_kria_ma, "is_file", lambda: False)():
            logger.warning("Pipeline MA: modelo Kria não disponível. Kria/Resposta podem não ser gerados. Verifique assets/templates.")
        resultado = mod.executar_pipeline_meio_ambiente_pdf(
            pdf_bytes,
            pasta_imagens=pasta_imagens,
            pasta_saida_kria=pasta_kria,
            pasta_saida_resp=pasta_resp,
            pasta_saida_eaf=pasta_eaf,
            pasta_saida_separar_nc=pasta_separar_nc,
            modelo_kria=modelo_kria_ma,
            modelo_resposta=modelo_resp_ma,
            modelo_kcor=None,
            pasta_saida_kcor=pasta_kcor,
            nome_origem=nome_origem,
        )
        buf = io.BytesIO()
        raiz_zip_ma = "Pipeline MA"
        adicionados = set()

        def _adicionar_arq(path, subpasta: str) -> None:
            if path is None:
                return
            p = path if isinstance(path, Path) else Path(str(path))
            if not p.is_file():
                return
            arcname = f"{raiz_zip_ma}/{subpasta}/{p.name}"
            if arcname in adicionados:
                return
            adicionados.add(arcname)
            try:
                zf.write(p, arcname)
            except Exception:
                pass

        with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
            # Fonte principal: conteúdo das pastas (onde o pipeline grava). Assim o ZIP reflete o que está em disco.
            pastas_zip = [
                (pasta_eaf, "EAF MA", ("*.xlsx",)),
                (pasta_kria, "Kria MA", ("*.xlsx",)),
                (pasta_resp, "Resposta MA", ("*.xlsx",)),
                (pasta_kcor, "Kcor-Kria Meio Ambiente", ("*.xlsx",)),
                (pasta_imagens, "Imagens MA", ("*.jpg", "*.jpeg", "*.png", "*.JPG", "*.JPEG", "*.PNG")),
            ]
            for folder, nome_pasta, extensoes in pastas_zip:
                try:
                    for ext in extensoes:
                        for f in folder.rglob(ext):
                            if f.is_file():
                                arcname = f"{raiz_zip_ma}/{nome_pasta}/{f.name}"
                                if arcname not in adicionados:
                                    adicionados.add(arcname)
                                    try:
                                        zf.write(str(f.resolve()), arcname)
                                    except Exception:
                                        try:
                                            zf.write(f, arcname)
                                        except Exception:
                                            pass
                except Exception:
                    pass
            # Incluir também os paths retornados pelo pipeline (caso estejam fora das pastas)
            _adicionar_arq(resultado.get("eaf"), "EAF MA")
            for path in resultado.get("kria") or []:
                _adicionar_arq(path, "Kria MA")
            for path in resultado.get("resposta") or []:
                _adicionar_arq(path, "Resposta MA")
            for path in resultado.get("kcor") or []:
                _adicionar_arq(path, "Kcor-Kria Meio Ambiente")
        zip_bytes = buf.getvalue()
        nome = "pipeline_ma.zip"
        ws.final.mkdir(parents=True, exist_ok=True)
        (ws.final / nome).write_bytes(zip_bytes)
        is_final = created or finalize
        if is_final:
            _update_job_json(ws, status="finished", stage="final", retain_hours=24 if created else 72)
            _purge_work_if_finished(ws)
        else:
            _update_job_json(ws, status="running", stage="stage2")
        if request.query_params.get("format") == "zip":
            return _stream_zip(zip_bytes, nome, ws.job_id)
        return JSONResponse(_nc_response(
            ws, "final" if is_final else "stage2",
            download_urls=[f"final/{nome}"],
            final_files=[nome] if is_final else None,
            step_label="Meio Ambiente",
            next_step_label="Acumulado",
        ))
    except HTTPException:
        raise
    except Exception as e:
        logger.error("nc/pipeline-meio-ambiente-pdf: %s", traceback.format_exc())
        raise HTTPException(500, str(e))


async def _inserir_nc(request, kria_zip, modelo_kcor, fotos_zip, modo, job_id: Optional[str] = None, finalize: bool = False, lote: Optional[str] = None):
    mod = _importar_modulo("inserir_nc_kria")
    lote_ok = (lote or "").strip() or None
    try:
        ws, created = resolve_workspace(job_id or None)
        work = ws.stage2 / "_work"
        work.mkdir(parents=True, exist_ok=True)

        pasta_kria = work / DIR_KRIA
        pasta_kria.mkdir(parents=True, exist_ok=True)
        with zipfile.ZipFile(io.BytesIO(_ler(kria_zip))) as zf:
            zf.extractall(str(pasta_kria))
        pasta_entrada = pasta_kria
        for sub in ("Kria", "kria"):
            cand = pasta_kria / sub
            if cand.is_dir() and list(cand.glob("*.xlsx")):
                pasta_entrada = cand
                break
        else:
            pasta_entrada = pasta_kria

        p_modelo = work / "modelo.xlsx"
        p_modelo.write_bytes(_ler(modelo_kcor) if modelo_kcor else _carregar_modelo_kcor(lote_ok))

        pasta_fotos = None
        if fotos_zip:
            pasta_fotos = work / DIR_IMAGENS_PDF
            pasta_fotos.mkdir(parents=True, exist_ok=True)
            with zipfile.ZipFile(io.BytesIO(_ler(fotos_zip))) as zf:
                zf.extractall(str(pasta_fotos))

        pasta_imagens = work / DIR_IMAGENS_CONSERVACAO
        pasta_saida   = work / DIR_CONSERVACAO
        pasta_imagens.mkdir(parents=True, exist_ok=True)
        pasta_saida.mkdir(parents=True, exist_ok=True)

        regime_artemig = modo == "conservacao" and (lote_ok or "").strip() == "50"
        if modo == "conservacao":
            mod.executar_conservacao(
                pasta_entrada=pasta_entrada,
                pasta_imagens=pasta_imagens,
                modelo_kcor=p_modelo,
                pasta_saida=pasta_saida,
                pasta_fotos_pdf=pasta_fotos,
                pasta_fotos_nc=pasta_fotos if fotos_zip else None,
                forcar_fallback=True,
                regime_artemig=regime_artemig,
            )
        else:
            mod.executar_meio_ambiente(
                pasta_entrada=pasta_entrada,
                pasta_imagens=pasta_imagens,
                modelo_kcor=p_modelo,
                pasta_saida=pasta_saida,
                pasta_fotos_pdf=pasta_fotos,
                pasta_fotos_nc=pasta_fotos if fotos_zip else None,
                forcar_fallback=True,
            )

        buf = io.BytesIO()
        pasta_zip_ma = "Kcor-Kria Meio Ambiente"  # pasta identificada dentro do ZIP de MA
        with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
            for f in pasta_saida.rglob("*.xlsx"):
                arcname = f"{pasta_zip_ma}/{f.name}" if modo == "meio_ambiente" else f.name
                zf.write(f, arcname)
        zip_bytes = buf.getvalue()
        nome = "kcor_conservacao.zip" if modo == "conservacao" else "kcor_ma.zip"
        ws.final.mkdir(parents=True, exist_ok=True)
        (ws.final / nome).write_bytes(zip_bytes)

        is_final = created or finalize
        if is_final:
            retain_hours = 24 if created else 72  # etapa isolada 24h, pipeline 72h
            _update_job_json(ws, status="finished", stage="final", retain_hours=retain_hours)
            _purge_work_if_finished(ws)
        else:
            _update_job_json(ws, status="running", stage="stage2")

        if request.query_params.get("format") == "zip":
            return _stream_zip(zip_bytes, nome, ws.job_id)
        step_label = "Conservação" if modo == "conservacao" else "Meio Ambiente"
        return JSONResponse(_nc_response(
            ws, "final" if is_final else "stage2",
            download_urls=[f"final/{nome}"],
            final_files=[nome] if is_final else None,
            step_label=step_label,
            next_step_label="Acumulado",
        ))
    except HTTPException:
        raise
    except Exception as e:
        logger.error("nc/inserir-%s: %s", modo, traceback.format_exc())
        raise HTTPException(500, str(e))


@router.post("/juntar", summary="M04 — Consolidar Acumulado")
async def nc_juntar(
    request: Request,
    kcor_zip: UploadFile = File(..., description="ZIP com Kcor-Kria individuais"),
    acumulado: Optional[UploadFile] = File(None, description="Planilha acumulada atual (opcional)"),
    nome_arquivo: Optional[str] = Form(None, description="Nome exato do arquivo de saída (ex.: 20260220-1313 - 20260220-1310 - 20260213 - CONSTATAÇÕES NC LOTE 13...xlsx)"),
    job_id: Optional[str] = Form(None),
    finalize: bool = Form(False),
):
    _check_auth(request)
    mod = _importar_modulo("juntar_arquivos")
    try:
        ws, created = resolve_workspace(job_id or None)
        work = ws.stage2 / "_work"
        work.mkdir(parents=True, exist_ok=True)

        pasta_kcor = work / DIR_KCOR_CONSERVACAO
        pasta_kcor.mkdir(parents=True, exist_ok=True)
        with zipfile.ZipFile(io.BytesIO(_ler(kcor_zip))) as zf:
            zf.extractall(str(pasta_kcor))

        p_acum = work / "acumulado_base.xlsx"
        if acumulado:
            p_acum.write_bytes(_ler(acumulado))
        else:
            mod.criar_base_acumulado(p_acum)

        pasta_saida = work / DIR_ACUMULADO
        pasta_saida.mkdir(parents=True, exist_ok=True)

        # Arquivos xlsx: busca recursiva (ZIP pode ter subpastas)
        xlsx_lista = [f for f in pasta_kcor.rglob("*.xlsx")
                      if not f.name.startswith("~")
                      and "Acumulado" not in f.name
                      and not f.name.startswith("_")]
        arquivos_entrada = sorted(xlsx_lista) if xlsx_lista else None

        resultado = mod.executar(
            pasta_entrada=pasta_kcor,
            arquivo_acumulado=p_acum,  # sempre pasta do job; se vazio, módulo cria base
            pasta_saida=pasta_saida,
            arquivos_entrada=arquivos_entrada,
            nome_arquivo_completo=nome_arquivo,
        )

        xlsx_bytes = None
        nome_final = "acumulado.xlsx"
        if resultado and Path(resultado).exists():
            xlsx_bytes = Path(resultado).read_bytes()
            nome_final = Path(resultado).name
        else:
            for f in pasta_saida.glob("*.xlsx"):
                xlsx_bytes = f.read_bytes()
                nome_final = f.name
                break
        if not xlsx_bytes:
            if arquivos_entrada:
                raise HTTPException(
                    400,
                    "Nenhum registro encontrado nos arquivos .xlsx do ZIP. "
                    "Verifique se as planilhas têm dados a partir da linha 2 (colunas A–Y). "
                    "Envie também o arquivo acumulado (relatório da rede) para consolidar.",
                )
            raise HTTPException(
                400,
                "Envie o arquivo acumulado (relatório da rede) para consolidar. Sem esse arquivo nada é gerado.",
            )
        ws.final.mkdir(parents=True, exist_ok=True)
        (ws.final / nome_final).write_bytes(xlsx_bytes)

        is_final = created or finalize
        if is_final:
            retain_hours = 24 if created else 72  # etapa isolada 24h, pipeline 72h
            _update_job_json(ws, status="finished", stage="final", retain_hours=retain_hours)
            _purge_work_if_finished(ws)
        else:
            _update_job_json(ws, status="running", stage="stage2")

        if request.query_params.get("format") == "xlsx":
            return _stream_xlsx(xlsx_bytes, nome_final, ws.job_id)
        return JSONResponse(_nc_response(
            ws, "final" if is_final else "stage2",
            download_urls=[f"final/{nome_final}"],
            final_files=[nome_final] if is_final else None,
            step_label="Acumulado",
            next_step_label="Inserir Nº, Calendário",
        ))
    except HTTPException:
        raise
    except Exception as e:
        logger.error("nc/juntar: %s", traceback.format_exc())
        raise HTTPException(500, str(e))


@router.post("/inserir-numero", summary="M05 — Inserir Nº Kria no Acumulado")
async def nc_inserir_numero(
    request: Request,
    acumulado: UploadFile = File(..., description="Planilha acumulada (.xlsx)"),
    numero_inicial: int = Form(1),
    sufixo: str = Form("26"),
    job_id: Optional[str] = Form(None),
    finalize: bool = Form(False),
):
    _check_auth(request)
    mod = _importar_modulo("inserir_numero_kria")
    try:
        ws, created = resolve_workspace(job_id or None)
        work = ws.stage2 / "_work"
        work.mkdir(parents=True, exist_ok=True)

        p = work / "acumulado.xlsx"
        p.write_bytes(_ler(acumulado))

        mod.executar(arquivo=p, numero_inicial=numero_inicial, sufixo=sufixo.strip())

        nome_arquivo = f"acumulado_{numero_inicial}{sufixo}.xlsx"
        xlsx_bytes = p.read_bytes()
        ws.final.mkdir(parents=True, exist_ok=True)
        (ws.final / nome_arquivo).write_bytes(xlsx_bytes)

        is_final = created or finalize
        if is_final:
            retain_hours = 24 if created else 72  # etapa isolada 24h, pipeline 72h
            _update_job_json(ws, status="finished", stage="final", retain_hours=retain_hours)
            _purge_work_if_finished(ws)
        else:
            _update_job_json(ws, status="running", stage="stage2")

        if request.query_params.get("format") == "xlsx":
            return _stream_xlsx(xlsx_bytes, nome_arquivo, ws.job_id)
        return JSONResponse(_nc_response(
            ws, "final" if is_final else "stage2",
            download_urls=[f"final/{nome_arquivo}"],
            final_files=[nome_arquivo] if is_final else None,
            step_label="Inserir Nº",
            next_step_label="—",
        ))
    except HTTPException:
        raise
    except Exception as e:
        logger.error("nc/inserir-numero: %s", traceback.format_exc())
        raise HTTPException(500, str(e))


@router.post("/exportar-calendario", summary="M06 — Exportar eventos para iCalendar (.ics)")
async def nc_exportar_calendario(
    request: Request,
    acumulado: UploadFile = File(..., description="Planilha acumulada com col Y preenchida (saída M05)"),
):
    """
    Gera um arquivo **.ics** (iCalendar) a partir da planilha acumulada.
    Cria um evento por NC com:
    - **Assunto:** TipoNC - Rodovia KM Sentido - Kria: {nº}
    - **Data:** extraída do campo Observações (col U)
    - **Descrição:** Obs Gestor + Data Constatação + Observações

    O arquivo pode ser importado diretamente no Outlook, Google Calendar ou Apple Calendar.
    Equivalente à macro `Art_06_EAF_Rot_Exportar_Calend` (modo .ics, sem Outlook).
    """
    _check_auth(request)
    mod = _importar_modulo("exportar_calendario")
    try:
        xlsx_bytes = _ler(acumulado)
        ics_bytes, n_eventos = await asyncio.to_thread(mod.gerar_ics_bytes, xlsx_bytes)
        stem = Path(acumulado.filename or "acumulado").stem
        return StreamingResponse(
            io.BytesIO(ics_bytes),
            media_type="text/calendar",
            headers={
                "Content-Disposition": f'attachment; filename="{stem}_eventos.ics"',
                "X-NC-Eventos": str(n_eventos),
            },
        )
    except HTTPException:
        raise
    except Exception as e:
        logger.error("nc/exportar-calendario: %s", traceback.format_exc())
        raise HTTPException(500, str(e))


@router.post("/organizar-imagens", summary="M08 — Organizar imagens por tipo de NC")
async def nc_organizar_imagens(
    request: Request,
    acumulado: UploadFile = File(..., description="Planilha acumulada (.xlsx) com colunas E, G, I, M, P, T, W, Y"),
    imagens_zip: UploadFile = File(..., description="ZIP com as imagens geradas no M03 (col W do acumulado)"),
):
    """
    Organiza as imagens em subpastas por tipo de NC.
    Nome de cada arquivo: `rodovia - sentido - km,metro - yyyymmdd - ddmmaaaa - evento.jpg`

    Estrutura do ZIP gerado:
    - `{Tipo NC}/` → imagens daquele tipo
    - `_Exportar/` → cópia extra dos tipos de pavimento (Depressão e Pano de Rolamento)

    Equivalente à macro `Salvar_IMG_NC_Artesp_Pasta_Sep`.
    """
    _check_auth(request)
    mod = _importar_modulo("salvar_imagem")
    try:
        xlsx_bytes = _ler(acumulado)
        zip_bytes  = _ler(imagens_zip)
        zip_saida, n_copiadas = await asyncio.to_thread(
            mod.organizar_imagens_bytes, xlsx_bytes, zip_bytes
        )
        return StreamingResponse(
            io.BytesIO(zip_saida),
            media_type="application/zip",
            headers={
                "Content-Disposition": 'attachment; filename="imagens_classificadas.zip"',
                "X-NC-Imagens": str(n_copiadas),
            },
        )
    except HTTPException:
        raise
    except Exception as e:
        logger.error("nc/organizar-imagens: %s", traceback.format_exc())
        raise HTTPException(500, str(e))


@router.post("/job", summary="Criar workspace por execução (job_id)")
async def nc_job_create(request: Request):
    """
    Cria um workspace por execução sob OUTPUT_PATH/nc/<job_id>/ com subpastas:
    input/, stage1/, stage2/, final/. Pipeline stateful: arquivos entre etapas
    sem re-upload. Retorna job_id para uso nos endpoints que aceitarem job_id.
    """
    _check_auth(request)
    ws = create_nc_workspace()
    return {
        "job_id": ws.job_id,
        "job_dir": str(ws.job_dir),
        "paths": {
            "input": str(ws.input),
            "stage1": str(ws.stage1),
            "stage2": str(ws.stage2),
            "final": str(ws.final),
        },
    }


@router.get("/job/{job_id}", summary="Info do workspace (job)")
async def nc_job_info(request: Request, job_id: str):
    """Retorna informações do workspace (existência, job.json com status/stages)."""
    _check_auth(request)
    ws = resolve_nc_workspace(job_id)
    if ws.job_dir.is_dir():
        _update_job_json(ws)
    exists = ws.job_dir.is_dir()
    subdirs = {}
    job_state = None
    if exists:
        for name in NC_SUBDIRS:
            p = getattr(ws, name)
            subdirs[name] = {"path": str(p), "exists": p.is_dir(), "files": _list_stage_files(p)}
        job_json = ws.job_dir / "job.json"
        if job_json.is_file():
            try:
                with open(job_json, "r", encoding="utf-8") as f:
                    job_state = json.load(f)
            except (json.JSONDecodeError, OSError):
                pass
    return {
        "job_id": ws.job_id,
        "job_dir": str(ws.job_dir),
        "exists": exists,
        "paths": {k: str(getattr(ws, k)) for k in NC_SUBDIRS},
        "subdirs": subdirs,
        "job": job_state,
    }


@router.get("/jobs/{job_id}", summary="Status do job com touch (job_manager)")
async def nc_job_status_touch(request: Request, job_id: str):
    """Carrega job via job_manager, atualiza last_access (touch) e devolve { ok, job }."""
    _check_auth(request)
    if not job_manager_carregar:
        raise HTTPException(501, detail="job_manager não disponível.")
    try:
        job = job_manager_carregar(job_id, touch=True)
    except HTTPException:
        raise
    return {"ok": True, "job": job}


@router.post("/start", summary="Etapa 1 — Upload único e extração (M01)")
async def nc_start(
    request: Request,
    arquivo: UploadFile = File(..., description="Planilha EAF (.xlsx ou .xls) — 1 arquivo"),
):
    """
    Recebe **1 arquivo** (EAF), salva em .../nc/<job_id>/input/, executa M01 (Separar NC)
    e grava intermediários em stage1/. Retorna job_id e links para stage1.
    Pipeline stateful: não é necessário re-enviar o arquivo nas próximas etapas.
    """
    _check_auth(request)
    mod = _importar_modulo("separar_nc")
    ws = create_nc_workspace()
    try:
        # Salvar em input/
        nome_safe = _safe_input_filename(arquivo.filename or "eaf.xlsx")
        input_path = ws.input / nome_safe
        input_path.write_bytes(_ler(arquivo))

        # M01 — Separar: EAF → XLS em stage1/
        arqs = mod.executar(input_path, pasta_destino=ws.stage1)
        if not arqs:
            raise HTTPException(500, detail="M01 não gerou arquivos.")

        # Zip dos XLS em stage1/nc_separados.zip
        zip_path = ws.stage1 / "nc_separados.zip"
        with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as zf:
            for a in arqs:
                p = Path(a)
                if p.exists():
                    zf.write(p, p.name)

        _update_job_json(ws, status="stage1")
        prefix = f"/outputs/nc/{ws.job_id}"
        return {
            "job_id": ws.job_id,
            "input_file": nome_safe,
            "stage1_files": ["nc_separados.zip"],
            "download_links": [f"{prefix}/stage1/nc_separados.zip"],
        }
    except HTTPException:
        raise
    except Exception as e:
        logger.error("nc/start: %s", traceback.format_exc())
        raise HTTPException(500, str(e))


class NCStage2Params(BaseModel):
    """Parâmetros da Etapa 2 (processamento a partir de stage1/)."""
    job_id: str
    numero_inicial: int = 1
    sufixo: str = "26"


@router.post("/stage2", summary="Etapa 2 — Processamento (M02→M06) com upload opcional de imagens PDF")
async def nc_stage2(
    request: Request,
    job_id: str = Form(..., description="ID do job (stage1 já executado)"),
    numero_inicial: int = Form(1, description="Número inicial para M04"),
    sufixo: str = Form("26", description="Sufixo para M04"),
    imagens_pdf_zip: Optional[UploadFile] = File(None, description="ZIP opcional com imagens PDF (N).jpg (saída do Extrair PDF). Se enviado, os e-mails embutirão as fotos."),
    lote: Optional[str] = Form(None, description="Lote 50 = ARTEMIG (templates em nc_artemig/assets/Template)"),
):
    """
    Recebe job_id + parâmetros e, opcionalmente, um ZIP com imagens extraídas do PDF.
    Lê stage1/, executa: M01 (já em stage1/) → Criar Email (Exportar + imagens em Imagens Provisórias - PDF) → M02→M03→M04→M05→M06.
    Se imagens_pdf_zip for enviado, é extraído em Imagens Provisórias - PDF antes do Criar Email, para os .eml incluírem as fotos.
    """
    _check_auth(request)
    ws = resolve_nc_workspace(job_id)
    if not ws.job_dir.is_dir():
        raise HTTPException(404, detail="Workspace não encontrado. Execute /nc/start primeiro.")
    _touch_job_access(ws)

    zip_stage1 = ws.stage1 / "nc_separados.zip"
    if not zip_stage1.is_file():
        raise HTTPException(400, detail="stage1/nc_separados.zip não encontrado. Execute /nc/start primeiro.")

    _update_job_json(ws, status="running")
    _garantir_path_nc()
    mod_modelo = _importar_modulo("gerar_modelo_foto")
    mod_inserir = _importar_modulo("inserir_nc_kria")
    mod_juntar = _importar_modulo("juntar_arquivos")
    mod_numero = _importar_modulo("inserir_numero_kria")
    mod_calendario = _importar_modulo("exportar_calendario")
    mod_criar_email = _importar_modulo("nc_criar_email")

    # Workspace persistente: zero temp. Tudo em stage2/_work/ (scratch) e stage2/ + final/ (saídas).
    ws.stage2.mkdir(parents=True, exist_ok=True)
    work = ws.stage2 / "_work"
    work.mkdir(exist_ok=True)
    pasta_xls = work / DIR_EXPORTAR
    pasta_xls.mkdir(exist_ok=True)
    try:
        with zipfile.ZipFile(zip_stage1, "r") as zf:
            zf.extractall(str(pasta_xls))

        # Imagens extraídas do PDF (opcional): para o e-mail embutir as fotos PDF (N).jpg
        pasta_fotos_pdf = work / DIR_IMAGENS_PDF
        if imagens_pdf_zip and imagens_pdf_zip.filename:
            pasta_fotos_pdf.mkdir(parents=True, exist_ok=True)
            _purge_dir_contents(pasta_fotos_pdf)
            with zipfile.ZipFile(io.BytesIO(await imagens_pdf_zip.read()), "r") as zf:
                zf.extractall(str(pasta_fotos_pdf))

        # Criar Email (macro NC_Artesp_Criar_Email): após M01 (Exportar), antes M02 — sequência das macros
        ws.final.mkdir(parents=True, exist_ok=True)
        pasta_emails = ws.final / "emails"
        pasta_emails.mkdir(exist_ok=True)
        try:
            mod_criar_email.executar(
                pasta_xls=pasta_xls,
                pasta_fotos_pdf=pasta_fotos_pdf if pasta_fotos_pdf.is_dir() else None,
                usar_outlook=False,
                pasta_saida_eml=pasta_emails,
            )
        except Exception as e_email:
            logger.warning("Módulo NC Email (após M01): %s", e_email)

        lote_ok = (lote or "").strip() or None
        p_kria = work / "modelo_kria.xlsx"
        p_resp = work / "modelo_resp.xlsx"
        p_kria.write_bytes(_carregar_modelo_kria(lote_ok))
        p_resp.write_bytes(_carregar_modelo_resp(lote_ok))

        pasta_kria = work / DIR_KRIA
        pasta_resp = work / DIR_RESPOSTAS_PENDENTES
        pasta_kria.mkdir(exist_ok=True)
        pasta_resp.mkdir(exist_ok=True)

        mod_modelo.executar(
            pasta_xls=pasta_xls,
            modelo_kria=p_kria,
            pasta_saida_kria=pasta_kria,
            modelo_resposta=p_resp,
            pasta_saida_resp=pasta_resp,
            pasta_fotos_nc=None,
            pasta_fotos_pdf=None,
        )
        kria_zip = ws.stage2 / "kria.zip"
        with zipfile.ZipFile(kria_zip, "w", zipfile.ZIP_DEFLATED) as zf:
            for f in pasta_kria.rglob("*.xlsx"):
                zf.write(f, f.name)

        p_modelo_kcor = work / "modelo_kcor.xlsx"
        p_modelo_kcor.write_bytes(_carregar_modelo_kcor(lote_ok))
        pasta_saida = work / DIR_CONSERVACAO
        pasta_saida.mkdir(exist_ok=True)
        pasta_imagens = work / DIR_IMAGENS_CONSERVACAO
        pasta_imagens.mkdir(exist_ok=True)
        mod_inserir.executar_conservacao(
            pasta_entrada=pasta_kria,
            pasta_imagens=pasta_imagens,
            modelo_kcor=p_modelo_kcor,
            pasta_saida=pasta_saida,
            pasta_fotos_pdf=None,
            forcar_fallback=True,
            regime_artemig=(lote_ok or "").strip() == "50",
        )
        kcor_zip = ws.stage2 / "kcor.zip"
        with zipfile.ZipFile(kcor_zip, "w", zipfile.ZIP_DEFLATED) as zf:
            for f in pasta_saida.rglob("*.xlsx"):
                zf.write(f, f.name)

        # M04 — Juntar (lê kcor extraído em work, grava acumulado em stage2)
        pasta_kcor_ext = work / DIR_KCOR_CONSERVACAO
        pasta_kcor_ext.mkdir(exist_ok=True)
        with zipfile.ZipFile(kcor_zip, "r") as zf:
            zf.extractall(str(pasta_kcor_ext))
        arquivos_kcor = [
            f for f in pasta_kcor_ext.rglob("*.xlsx")
            if not f.name.startswith("~") and "Acumulado" not in f.name and not f.name.startswith("_")
        ]
        arquivos_kcor = sorted(arquivos_kcor) if arquivos_kcor else None
        # Base do acumulado dentro do job (evita usar M04_ACUMULADO do config, que pode não existir)
        p_acum_base = work / "acumulado_base.xlsx"
        path_acumulado = mod_juntar.executar(
            pasta_entrada=None,
            arquivo_acumulado=p_acum_base,
            pasta_saida=ws.stage2,
            arquivos_entrada=arquivos_kcor,
        )
        if not path_acumulado or not Path(path_acumulado).exists():
            raise HTTPException(500, detail="M04 não gerou acumulado.")
        acumulado_path = Path(path_acumulado)

        # M05 — Inserir número (modifica in-place; depois copiamos para numerado)
        mod_numero.executar(
            acumulado_path,
            numero_inicial,
            sufixo=sufixo,
        )
        numerado_path = ws.stage2 / "acumulado_numerado.xlsx"
        shutil.copy2(str(acumulado_path), str(numerado_path))

        # M06 — Exportar calendário (.ics) em final/
        ws.final.mkdir(parents=True, exist_ok=True)
        mod_calendario.executar(
            numerado_path,
            usar_outlook=False,
            pasta_saida_ics=ws.final,
            executar_mod08=False,
        )
        shutil.copy2(str(numerado_path), str(ws.final / "acumulado_numerado.xlsx"))

        # ZIP final único (estratégia 5): um arquivo para auditoria; inclui .xlsx, .ics e pasta emails/
        zip_final_name = f"nc_{ws.job_id}_artesp.zip"
        zip_final_path = ws.final / zip_final_name
        with zipfile.ZipFile(zip_final_path, "w", zipfile.ZIP_DEFLATED) as zf:
            for f in ws.final.iterdir():
                if f.is_file() and f.suffix.lower() in (".xlsx", ".ics"):
                    zf.write(f, f.name)
            pasta_emails = ws.final / "emails"
            if pasta_emails.is_dir():
                for eml in pasta_emails.rglob("*"):
                    if eml.is_file():
                        zf.write(eml, eml.relative_to(ws.final))
        # Remove arquivos soltos em final/ (fica só o ZIP)
        for f in list(ws.final.iterdir()):
            if f.is_file() and f != zip_final_path:
                try:
                    f.unlink()
                except OSError:
                    pass

        # Apagar intermediários cedo (estratégia 2): sucesso → só final/ obrigatório
        _purge_dir_contents(ws.stage1)
        _purge_dir_contents(ws.stage2)

        _update_job_json(
            ws,
            status="finished",
            log_summary={"errors": 0, "warnings": 0},
            retain_hours=72.0,
        )
        prefix = f"/outputs/nc/{ws.job_id}"
        return {
            "job_id": ws.job_id,
            "download_links": [f"{prefix}/final/{zip_final_name}"],
            "download_zip": zip_final_name,
        }
    except HTTPException:
        raise
    except Exception as e:
        logger.error("nc/stage2: %s", traceback.format_exc())
        try:
            _update_job_json(
                ws,
                status="failed",
                log_summary={"errors": 1, "warnings": 0, "message": str(e)[:200]},
                retain_hours=24.0,
            )
        except Exception:
            pass
        raise HTTPException(500, str(e))


@router.get("/", summary="Status do módulo NC Artesp")
async def nc_info():
    return {
        "modulo": "nc_artesp",
        "pipeline": "M01 → Criar Email → M02 → M03 → M04 → M05 → M06 → M08",
        "nc_proj_disponivel": _nc_proj_disponivel(),
        "nc_proj_path": str(_NC_PROJ),
        "nc_output_path": str(_nc_output_path()),
        "endpoints": [
            "POST /nc/start                  → Etapa 1: upload 1 arquivo → input/ + stage1/ (job_id)",
            "POST /nc/stage2                 → Etapa 2: job_id + params [opcional: imagens_pdf_zip] → stage2/ + final/ (links)",
            "GET  /outputs/nc/{job_id}/{subpath} → Download arquivo do job",
            "POST /nc/job                    → Criar workspace vazio (job_id)",
            "GET  /nc/job/{job_id}           → Info do workspace",
            "POST /nc/extrair-pdf             → PDF NC → ZIP nc(N).jpg + PDF(N).jpg",
            "POST /nc/analisar-pdf           → PDF NC → PDF análise (gaps KM, emergenciais, por tipo)",
            "POST /nc/separar                → M01: EAF → ZIP XLS individuais",
            "POST /nc/criar-email             → ZIP XLS + opcional imagens PDF → ZIP .eml",
            "POST /nc/stage2                 → M01+Email+M02→M06 (Email com fotos se enviar imagens_pdf_zip)",
            "POST /nc/gerar-modelo-foto      → M02: XLS ZIP → ZIP Kria + Resposta",
            "POST /nc/inserir-conservacao    → M03: Kria → ZIP Kcor-Kria Conservação",
            "POST /nc/inserir-meio-ambiente  → M07: Kria MA → ZIP Kcor-Kria MA",
            "POST /nc/juntar                 → M04: Kcor → XLSX Acumulado",
            "POST /nc/inserir-numero         → M05: Acumulado numerado",
            "POST /nc/exportar-calendario    → M06: Acumulado → .ics (iCalendar)",
            "POST /nc/organizar-imagens      → M08: Acumulado + ZIP imagens → ZIP classificado por tipo",
        ],
    }
