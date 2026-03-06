"""
render_api/job_manager.py
────────────────────────────────────────────────────────────────────────────
Job manager para pipeline NC ARTESP: workspace por job + job.json template.
Escrita atômica (tmp + os.replace), retenção por estado, limpeza com flags.pinned.
Uso: criar_job_nc(), carregar_job_nc(), salvar_job_nc(), registrar_arquivo_job(),
     finalizar_job(), limpar_jobs_nc().
"""

from __future__ import annotations

import datetime
import json
import os
import secrets
import shutil
from collections import defaultdict
from pathlib import Path
from typing import Any, Dict, Optional

from fastapi import HTTPException

# Raiz dos jobs NC (OUTPUT_PATH/nc) — mesmo critério do app
_env = (os.getenv("ARTESP_OUTPUT_DIR") or "").strip()
if _env:
    NC_JOBS_ROOT = (Path(_env).resolve() / "nc").resolve()
else:
    import platform
    if platform.system() == "Linux":
        NC_JOBS_ROOT = (Path("/data/outputs") / "nc").resolve()
    else:
        NC_JOBS_ROOT = (Path(__file__).resolve().parent.parent / "outputs" / "nc").resolve()
NC_JOBS_ROOT.mkdir(parents=True, exist_ok=True)


def _agora_iso_tz() -> str:
    return datetime.datetime.now().astimezone().isoformat(timespec="seconds")


def _safe_job_id(job_id: str) -> str:
    job_id = (job_id or "").strip()
    if not job_id:
        raise HTTPException(status_code=400, detail="job_id ausente")
    if any(x in job_id for x in ("..", "/", "\\", "\x00")):
        raise HTTPException(status_code=400, detail="job_id inválido")
    return job_id


def _job_dir(job_id: str) -> Path:
    job_id = _safe_job_id(job_id)
    d = (NC_JOBS_ROOT / job_id).resolve()
    try:
        d.relative_to(NC_JOBS_ROOT.resolve())
    except ValueError:
        raise HTTPException(status_code=400, detail="Acesso negado (job_dir).")
    return d


def _job_file(job_id: str) -> Path:
    return _job_dir(job_id) / "job.json"


def _template_job(job_id: str, params: Optional[Dict[str, Any]] = None, retain_hours: int = 72) -> Dict[str, Any]:
    now = _agora_iso_tz()
    retain_until = (
        datetime.datetime.now().astimezone() + datetime.timedelta(hours=max(1, retain_hours))
    ).isoformat(timespec="seconds")
    return {
        "job_id": job_id,
        "pipeline": "nc_artesp",
        "status": "running",
        "stage": "stage1",
        "created_at": now,
        "updated_at": now,
        "last_access": now,
        "retain_until": retain_until,
        "params": params or {},
        "files": {"input": [], "stage1": [], "stage2": [], "final": []},
        "stats": {"bytes_written": 0, "n_rows": None, "n_items": None},
        "errors": [],
        "warnings": [],
        "flags": {"pinned": False, "keep_intermediates": False},
    }


def criar_job_nc(
    params: Optional[Dict[str, Any]] = None,
    retain_hours: int = 72,
) -> Dict[str, Any]:
    """Cria workspace do job + job.json (template completo)."""
    job_id = f"nc_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}_{secrets.token_urlsafe(6)}"
    d = _job_dir(job_id)
    d.mkdir(parents=True, exist_ok=True)
    for sub in ("input", "stage1", "stage2", "final"):
        (d / sub).mkdir(parents=True, exist_ok=True)
    job = _template_job(job_id, params=params, retain_hours=retain_hours)
    salvar_job_nc(job)
    return job


def carregar_job_nc(job_id: str, touch: bool = True) -> Dict[str, Any]:
    """Lê job.json. Se touch=True, atualiza last_access e updated_at."""
    p = _job_file(job_id)
    if not p.is_file():
        raise HTTPException(status_code=404, detail="Job não encontrado.")
    try:
        with open(p, "r", encoding="utf-8") as f:
            job = json.load(f)
    except Exception:
        raise HTTPException(status_code=500, detail="Falha ao ler job.json.")
    # Garante estrutura mínima (compatível com job.json escrito por nc_router)
    for key, default in (("params", {}), ("files", {}), ("stats", {}), ("errors", []), ("warnings", []), ("flags", {})):
        job.setdefault(key, default)
    job.setdefault("files", {}).setdefault("input", [])
    job.setdefault("files", {}).setdefault("stage1", [])
    job.setdefault("files", {}).setdefault("stage2", [])
    job.setdefault("files", {}).setdefault("final", [])
    if touch:
        job["last_access"] = _agora_iso_tz()
        job["updated_at"] = job["last_access"]
        salvar_job_nc(job)
    return job


def salvar_job_nc(job: Dict[str, Any]) -> None:
    """Escrita atômica do job.json (job.json.tmp -> job.json)."""
    job_id = _safe_job_id(job.get("job_id", ""))
    d = _job_dir(job_id)
    p = d / "job.json"
    tmp = d / "job.json.tmp"
    now = _agora_iso_tz()
    job.setdefault("created_at", now)
    job.setdefault("updated_at", now)
    job.setdefault("last_access", now)
    try:
        with open(tmp, "w", encoding="utf-8") as f:
            json.dump(job, f, ensure_ascii=False, indent=2)
            f.flush()
            os.fsync(f.fileno())
        os.replace(tmp, p)
    except Exception as e:
        try:
            if tmp.exists():
                tmp.unlink()
        except Exception:
            pass
        raise HTTPException(status_code=500, detail=f"Falha ao salvar job.json: {e}")


def registrar_arquivo_job(job: Dict[str, Any], etapa: str, nome_arquivo: str) -> Dict[str, Any]:
    """Registra arquivo gerado em job['files'][etapa] e atualiza timestamps."""
    etapa = (etapa or "").strip()
    if etapa not in ("input", "stage1", "stage2", "final"):
        raise HTTPException(status_code=400, detail="Etapa inválida.")
    nome_arquivo = os.path.basename(nome_arquivo or "").strip()
    if not nome_arquivo:
        return job
    files = job.setdefault("files", {"input": [], "stage1": [], "stage2": [], "final": []})
    files.setdefault(etapa, [])
    if nome_arquivo not in files[etapa]:
        files[etapa].append(nome_arquivo)
    job["updated_at"] = _agora_iso_tz()
    job["last_access"] = job["updated_at"]
    salvar_job_nc(job)
    return job


def finalizar_job(
    job_id: str,
    ok: bool,
    keep_intermediates: Optional[bool] = None,
) -> Dict[str, Any]:
    """Marca job finished/failed e opcionalmente remove stage1/ e stage2/."""
    job = carregar_job_nc(job_id, touch=True)
    job["status"] = "finished" if ok else "failed"
    job["stage"] = "final"
    if keep_intermediates is not None:
        job.setdefault("flags", {})["keep_intermediates"] = bool(keep_intermediates)
    job["updated_at"] = _agora_iso_tz()
    job["last_access"] = job["updated_at"]
    salvar_job_nc(job)
    if ok and not job.get("flags", {}).get("keep_intermediates", False):
        d = _job_dir(job_id)
        for sub in ("stage1", "stage2"):
            try:
                sub_path = d / sub
                if sub_path.is_dir():
                    shutil.rmtree(sub_path, ignore_errors=True)
            except Exception:
                pass
    return job


def _parse_iso(dt_str: str) -> Optional[datetime.datetime]:
    if not dt_str:
        return None
    try:
        return datetime.datetime.fromisoformat(dt_str.replace("Z", "+00:00"))
    except Exception:
        return None


def _last_access_to_datetime(job: Dict[str, Any]) -> Optional[datetime.datetime]:
    """last_access pode ser ISO ou timestamp numérico (compat nc_router)."""
    la = job.get("last_access")
    if la is None:
        return _parse_iso(job.get("updated_at") or "") or _parse_iso(job.get("created_at") or "")
    if isinstance(la, (int, float)):
        try:
            return datetime.datetime.fromtimestamp(la, tz=datetime.timezone.utc)
        except Exception:
            return None
    return _parse_iso(str(la))


def limpar_jobs_nc(
    retention_finished_h: int = 72,
    retention_failed_h: int = 24,
    retention_archived_h: int = 168,
    retention_orphan_h: int = 24,
    dry_run: bool = False,
) -> Dict[str, Any]:
    """
    Limpa jobs antigos em NC_JOBS_ROOT por status, last_access, retain_until, flags.pinned.
    job.json corrompido (não parseia) → tratado como órfão, retenção retention_orphan_h (ex.: 24h).
    Retorna: removidos, ignorados, bytes_liberados, removidos_por_status.
    """
    now = datetime.datetime.now().astimezone()
    now_ts = now.timestamp()
    removidos = 0
    ignorados = 0
    bytes_liberados = 0
    removidos_por_status: Dict[str, int] = defaultdict(int)
    if not NC_JOBS_ROOT.is_dir():
        return {
            "removidos": 0,
            "ignorados": 0,
            "bytes_liberados": 0,
            "removidos_por_status": {},
        }
    for d in list(NC_JOBS_ROOT.iterdir()):
        if not d.is_dir():
            continue
        job_json = d / "job.json"
        if not job_json.is_file():
            ignorados += 1
            continue
        try:
            job = json.loads(job_json.read_text(encoding="utf-8"))
        except (json.JSONDecodeError, OSError):
            # job.json corrompido → órfão: remover após retention_orphan_h pelo mtime
            try:
                mtime = job_json.stat().st_mtime
                idade_h = (now_ts - mtime) / 3600.0
            except Exception:
                ignorados += 1
                continue
            if idade_h < max(1, retention_orphan_h):
                ignorados += 1
                continue
            try:
                for p in d.rglob("*"):
                    if p.is_file():
                        bytes_liberados += p.stat().st_size
            except Exception:
                pass
            if not dry_run:
                try:
                    shutil.rmtree(d)
                except Exception:
                    ignorados += 1
                    continue
            removidos += 1
            removidos_por_status["orphan"] += 1
            continue
        flags = job.get("flags") or {}
        if flags.get("pinned"):
            ignorados += 1
            continue
        status = (job.get("status") or "running").strip().lower()
        last_access = _last_access_to_datetime(job)
        retain_until = _parse_iso(job.get("retain_until") or "")
        if retain_until and retain_until > now:
            ignorados += 1
            continue
        retain_ts = job.get("retain_until_ts")
        if retain_ts is not None and now_ts <= retain_ts:
            ignorados += 1
            continue
        if last_access is None:
            last_access = now
        idade_h = (now - last_access).total_seconds() / 3600.0
        limite = None
        if status == "finished":
            limite = retention_finished_h
        elif status == "failed":
            limite = retention_failed_h
        elif status == "archived":
            limite = retention_archived_h
        else:
            # running e qualquer outro status: nunca removido pela limpeza
            ignorados += 1
            continue
        if idade_h >= max(1, limite):
            try:
                for p in d.rglob("*"):
                    if p.is_file():
                        bytes_liberados += p.stat().st_size
            except Exception:
                pass
            if not dry_run:
                try:
                    shutil.rmtree(d)
                except Exception:
                    ignorados += 1
                    continue
            removidos += 1
            removidos_por_status[status] += 1
        else:
            ignorados += 1
    return {
        "removidos": removidos,
        "ignorados": ignorados,
        "bytes_liberados": bytes_liberados,
        "removidos_por_status": dict(removidos_por_status),
    }
