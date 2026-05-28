from __future__ import annotations

import re
import shutil
import sys
import threading
import uuid
from datetime import datetime
from pathlib import Path
from typing import Any, Literal, Optional

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from fastapi import FastAPI, File, HTTPException, Query, UploadFile
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse, HTMLResponse
from fastapi.staticfiles import StaticFiles
from pydantic import BaseModel, Field

from backend.excel_io import workbook_from_json, workbook_to_json
from backend import tls_bot_service
from prompt_templates import get_prompt, reset_prompt, save_prompt
from backend.storage import JobStore

BASE_DIR = Path(__file__).resolve().parent
UPLOAD_DIR = BASE_DIR / "uploads"
OUTPUT_DIR = BASE_DIR / "outputs"
DATA_DIR = BASE_DIR / "data"
VERSIONS_DIR = DATA_DIR / "versions"

UPLOAD_DIR.mkdir(parents=True, exist_ok=True)
OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

store = JobStore(DATA_DIR, VERSIONS_DIR)

app = FastAPI(title="TC Generator API", version="2.0.0")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_methods=["*"],
    allow_headers=["*"],
)

PipelineMode = Literal[
    "all",
    "e2e",
    "e2e_max",
    "e2e_short",
    "output",
    "output_short",
    "multi_responses",
]
DataGenMode = Literal["llm", "tls", "hybrid", "asr_noise"]

MODES_ALL = [
    ("e2e_short", "Test case"),
    ("output_short", "Test Output"),
    ("multi_responses", "Test case Đa Thoại"),
]

MODE_TO_SHEET: dict[str, str] = {
    "e2e": "Test case",
    "e2e_max": "Test case",
    "e2e_short": "Test case",
    "output": "Test Output",
    "output_short": "Test Output",
    "multi_responses": "Test case Đa Thoại",
}


class RunRequest(BaseModel):
    file_id: str
    mode: PipelineMode = "all"
    sheet: Optional[str] = None
    root: Optional[str] = None
    max_depth: int = 200
    gen_data: bool = False
    apply_asr_noise: bool = False


TcPipelineMode = Literal["e2e_output", "multi_responses"]


class RunDataRequest(BaseModel):
    file_id: str
    data_mode: DataGenMode = "hybrid"
    tc_pipeline_mode: TcPipelineMode = "e2e_output"


class PatchJobRequest(BaseModel):
    download_filename: Optional[str] = None


class WorkbookPayload(BaseModel):
    sheet_names: list[str] = Field(default_factory=list)
    active_sheet: str = "Sheet1"
    rows: list[list[str]] = Field(default_factory=list)
    save_as_filename: Optional[str] = None


def _safe_download_name(name: str) -> str:
    cleaned = re.sub(r'[<>:"/\\|?*]', "_", name.strip())
    if not cleaned:
        raise HTTPException(400, "Invalid filename")
    if not cleaned.lower().endswith((".xlsx", ".xlsm", ".xls")):
        cleaned += ".xlsx"
    return cleaned


def _find_upload(file_id: str) -> Path:
    for ext in (".xlsx", ".xlsm", ".xls"):
        p = UPLOAD_DIR / f"{file_id}{ext}"
        if p.exists():
            return p
    raise FileNotFoundError(f"Uploaded file not found: {file_id!r}")


def _upload_meta_path(file_id: str) -> Path:
    return UPLOAD_DIR / f"{file_id}.meta.json"


def _read_upload_meta(file_id: str) -> dict[str, Any]:
    import json

    p = _upload_meta_path(file_id)
    if p.exists():
        return json.loads(p.read_text(encoding="utf-8"))
    return {"filename": f"{file_id}.xlsx"}


def _write_upload_meta(file_id: str, meta: dict[str, Any]) -> None:
    import json

    _upload_meta_path(file_id).write_text(
        json.dumps(meta, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )


def _asr_noise_progress_fn(job_id: str, base_pct: int, end_pct: int):
    """Return a progress callback that maps row progress to [base_pct, end_pct]."""
    def _fn(done: int, total: int) -> None:
        if total > 0:
            pct = base_pct + int((end_pct - base_pct) * done / total)
            store.update_job(
                job_id,
                progress=f"Applying ASR noise… {done}/{total}",
                progress_pct=min(pct, end_pct),
            )
    return _fn


def _run_pipeline_job(job_id: str, req: RunRequest) -> None:
    try:
        store.update_job(job_id, status="running", progress="Locating uploaded file…", progress_pct=0)
        excel_path = _find_upload(req.file_id)

        from pipeline_tc import run_pipeline
        from excel_to_json import read_data_schema_sheet
        from tc_to_excel import _add_checklist_sheet, _add_data_input_sheet, _add_data_schema_sheet
        import openpyxl as xl

        job_dir = OUTPUT_DIR / job_id
        job_dir.mkdir(exist_ok=True)

        sheet = int(req.sheet) if req.sheet and req.sheet.isdigit() else req.sheet

        if req.mode == "all":
            # With ASR noise: pipelines share 0-65%, save=70%, noise=70-95%, done=100%
            # Without ASR noise: pipelines share 0-80%, save=90%, done=100%
            pipeline_end = 65 if req.apply_asr_noise else 80
            n = len(MODES_ALL)

            schema_keys, schema_value_rows, schema_all_rows = read_data_schema_sheet(excel_path)
            combined_wb = xl.Workbook()
            combined_wb.remove(combined_wb.active)
            _add_checklist_sheet(combined_wb)

            results = []
            for i, (mode, tc_sheet_name) in enumerate(MODES_ALL):
                pct_start = int(pipeline_end * i / n)
                store.update_job(job_id, progress=f"Running {mode}… ({i + 1}/{n})", progress_pct=pct_start)
                _, _, _, count, _ = run_pipeline(
                    mode=mode,
                    excel_path=excel_path,
                    sheet=sheet,
                    root=req.root,
                    max_depth=req.max_depth,
                    rows_out=job_dir / f"rows_{mode}.json",
                    testcases_out=job_dir / f"testcases_{mode}.json",
                    excel_out=job_dir / f"testcases_{mode}.xlsx",
                    gen_data=req.gen_data,
                    combined_wb=combined_wb,
                    tc_sheet_name=tc_sheet_name,
                )
                results.append({"mode": mode, "count": count, "sheet": tc_sheet_name})
                pct_done = int(pipeline_end * (i + 1) / n)
                store.update_job(job_id, progress_pct=pct_done)

            save_pct = 70 if req.apply_asr_noise else 90
            store.update_job(job_id, progress="Saving combined workbook…", progress_pct=save_pct)
            _add_data_input_sheet(
                combined_wb,
                schema_keys=schema_keys or None,
                schema_value_rows=schema_value_rows or None,
            )
            _add_data_schema_sheet(combined_wb, schema_all_rows=schema_all_rows or None)

            output_excel = job_dir / "testcases_all.xlsx"
            combined_wb.save(output_excel)
            output_name = "testcases_all.xlsx"

            if req.apply_asr_noise:
                from generate_asr_noise import fill_excel_with_asr_noise
                store.update_job(job_id, progress="Applying ASR noise…", progress_pct=70)
                fill_excel_with_asr_noise(
                    output_excel, output_excel,
                    progress_fn=_asr_noise_progress_fn(job_id, 70, 95),
                )

            store.add_version(
                job_id,
                kind="output",
                src_path=output_excel,
                filename=output_name,
            )

            store.update_job(
                job_id,
                status="completed",
                progress="Done",
                progress_pct=100,
                completed_at=datetime.utcnow().isoformat(),
                download_filename=output_name,
                result={
                    "modes": results,
                    "total": sum(r["count"] for r in results),
                    "output_file": output_name,
                },
            )
        else:
            # Individual mode: running=10%, pipeline done=70% (or 85% without noise), ASR 70-95%, done=100%
            pipeline_done_pct = 70 if req.apply_asr_noise else 85
            tc_sheet_name = MODE_TO_SHEET[req.mode]
            store.update_job(job_id, progress=f"Running {req.mode}…", progress_pct=10)
            _, _, excel_out, count, _ = run_pipeline(
                mode=req.mode,
                excel_path=excel_path,
                sheet=sheet,
                root=req.root,
                max_depth=req.max_depth,
                rows_out=job_dir / f"rows_{req.mode}.json",
                testcases_out=job_dir / f"testcases_{req.mode}.json",
                excel_out=job_dir / f"testcases_{req.mode}.xlsx",
                gen_data=req.gen_data,
            )
            store.update_job(job_id, progress_pct=pipeline_done_pct)
            output_name = f"testcases_{req.mode}.xlsx"

            if req.apply_asr_noise:
                from generate_asr_noise import fill_excel_with_asr_noise
                store.update_job(job_id, progress="Applying ASR noise…", progress_pct=70)
                fill_excel_with_asr_noise(
                    excel_out, excel_out,
                    progress_fn=_asr_noise_progress_fn(job_id, 70, 95),
                )

            store.add_version(
                job_id,
                kind="output",
                src_path=excel_out,
                filename=output_name,
            )
            store.update_job(
                job_id,
                status="completed",
                progress="Done",
                progress_pct=100,
                completed_at=datetime.utcnow().isoformat(),
                download_filename=output_name,
                result={
                    "modes": [{"mode": req.mode, "count": count, "sheet": tc_sheet_name}],
                    "total": count,
                    "output_file": output_name,
                },
            )

    except Exception as exc:
        store.update_job(
            job_id,
            status="failed",
            progress="",
            completed_at=datetime.utcnow().isoformat(),
            error=str(exc),
        )


def _run_data_gen_job(job_id: str, req: RunDataRequest) -> None:
    try:
        store.update_job(job_id, status="running", progress="Reading testcase workbook…", progress_pct=5)
        excel_path = _find_upload(req.file_id)
        job_dir = OUTPUT_DIR / job_id
        job_dir.mkdir(exist_ok=True)

        is_multi = req.tc_pipeline_mode == "multi_responses"
        output_name = (
            f"testdata_multi_{req.data_mode}.xlsx" if is_multi
            else f"testdata_{req.data_mode}.xlsx"
        )
        out_path = job_dir / output_name

        if is_multi:
            # Multi-response format: read tc_to_excel multi_responses output,
            # fill 'User Message' column with Hybrid-generated data,
            # consecutive same-scenario TCs share the same generated data.
            from generate_test_data_hybrid import fill_excel_multi_from_excel

            store.update_job(
                job_id,
                progress="Generating test data (Multi-response, Hybrid)…",
                progress_pct=10,
            )
            fill_excel_multi_from_excel(excel_path, out_path)
            store.update_job(job_id, progress_pct=90)

        elif req.data_mode == "llm":
            from generate_test_data_llm import fill_excel_from_excel

            store.update_job(job_id, progress="Generating test data with LLM…", progress_pct=10)
            fill_excel_from_excel(excel_path, out_path)
            store.update_job(job_id, progress_pct=90)
        elif req.data_mode == "tls":
            from generate_test_data_neuron import fill_excel_from_excel

            store.update_job(job_id, progress="Generating test data with TLS Bot…", progress_pct=10)
            fill_excel_from_excel(excel_path, out_path)
            store.update_job(job_id, progress_pct=90)
        elif req.data_mode == "asr_noise":
            from generate_asr_noise import fill_excel_with_asr_noise

            store.update_job(job_id, progress="Applying ASR noise to test data…", progress_pct=10)
            fill_excel_with_asr_noise(
                excel_path, out_path,
                progress_fn=_asr_noise_progress_fn(job_id, 10, 95),
            )
        else:
            from generate_test_data_hybrid import fill_excel_from_excel

            store.update_job(job_id, progress="Generating test data with Hybrid (TLS + LLM)…", progress_pct=10)
            fill_excel_from_excel(excel_path, out_path)
            store.update_job(job_id, progress_pct=90)

        store.add_version(
            job_id,
            kind="output",
            src_path=out_path,
            filename=output_name,
        )

        mode_label = f"multi_responses/{req.data_mode}" if is_multi else req.data_mode
        store.update_job(
            job_id,
            status="completed",
            progress="Done",
            progress_pct=100,
            completed_at=datetime.utcnow().isoformat(),
            download_filename=output_name,
            result={
                "modes": [{"mode": mode_label, "count": 0, "sheet": "Test case Đa Thoại" if is_multi else "Test Data"}],
                "total": 0,
                "output_file": output_name,
                "data_mode": req.data_mode,
                "tc_pipeline_mode": req.tc_pipeline_mode,
            },
        )
    except Exception as exc:
        store.update_job(
            job_id,
            status="failed",
            progress="",
            completed_at=datetime.utcnow().isoformat(),
            error=str(exc),
        )


def _resolve_output_path(job: dict[str, Any]) -> Path:
    job_id = job["job_id"]
    latest = store.latest_version_path(job_id, "output") or store.latest_version_path(
        job_id, "edited_output"
    )
    if latest:
        return latest
    result = job.get("result") or {}
    output_file = result.get("output_file", "output.xlsx")
    return OUTPUT_DIR / job_id / output_file


# ── Health ─────────────────────────────────────────────────────

@app.get("/health/live")
def liveness():
    return {"status": "ok"}


@app.get("/health/ready")
def readiness():
    return {"status": "ok", "jobs": len(store.list_jobs(limit=10000))}


# ── Upload ─────────────────────────────────────────────────────

@app.post("/api/upload")
async def upload_file(file: UploadFile = File(...)):
    if not file.filename:
        raise HTTPException(400, "No filename provided")
    suffix = Path(file.filename).suffix.lower()
    if suffix not in {".xlsx", ".xlsm", ".xls"}:
        raise HTTPException(400, "Only Excel files (.xlsx .xlsm .xls) are accepted")

    file_id = str(uuid.uuid4())
    dest = UPLOAD_DIR / f"{file_id}{suffix}"
    content = await file.read()
    dest.write_bytes(content)

    _write_upload_meta(
        file_id,
        {"filename": file.filename, "uploaded_at": datetime.utcnow().isoformat()},
    )

    return {"file_id": file_id, "filename": file.filename, "size": dest.stat().st_size}


# ── Source workbook (edit input) ───────────────────────────────

@app.get("/api/files/{file_id}/workbook")
def get_source_workbook(file_id: str, sheet: Optional[str] = None):
    try:
        path = _find_upload(file_id)
    except FileNotFoundError:
        raise HTTPException(404, "File not found")
    sheet_arg: str | int | None = int(sheet) if sheet and sheet.isdigit() else sheet
    return workbook_to_json(path, sheet=sheet_arg)


@app.put("/api/files/{file_id}/workbook")
def save_source_workbook(file_id: str, payload: WorkbookPayload):
    try:
        path = _find_upload(file_id)
    except FileNotFoundError:
        raise HTTPException(404, "File not found")

    meta = _read_upload_meta(file_id)
    filename = payload.save_as_filename or meta.get("filename", path.name)
    filename = _safe_download_name(filename)

    temp = UPLOAD_DIR / f"{file_id}_edit_{uuid.uuid4().hex[:8]}.xlsx"
    workbook_from_json(
        sheet_names=payload.sheet_names or [payload.active_sheet],
        active_sheet=payload.active_sheet,
        rows=payload.rows,
        dest=temp,
    )
    shutil.copy2(temp, path)
    temp.unlink(missing_ok=True)

    meta["filename"] = filename
    meta["last_edited_at"] = datetime.utcnow().isoformat()
    _write_upload_meta(file_id, meta)

    return {
        "file_id": file_id,
        "filename": filename,
        "message": "Source workbook saved",
    }


# ── Pipeline ───────────────────────────────────────────────────

@app.post("/api/run")
def start_pipeline(req: RunRequest):
    try:
        excel_path = _find_upload(req.file_id)
    except FileNotFoundError:
        raise HTTPException(404, f"Uploaded file not found: {req.file_id!r}")

    meta = _read_upload_meta(req.file_id)
    job_id = store.create_job(
        job_type="pipeline",
        file_id=req.file_id,
        source_filename=meta.get("filename", excel_path.name),
        mode=req.mode,
    )
    store.add_version(
        job_id,
        kind="source",
        src_path=excel_path,
        filename=meta.get("filename", excel_path.name),
    )

    threading.Thread(target=_run_pipeline_job, args=(job_id, req), daemon=True).start()
    return {"job_id": job_id}


@app.post("/api/run-data")
def start_data_generator(req: RunDataRequest):
    try:
        excel_path = _find_upload(req.file_id)
    except FileNotFoundError:
        raise HTTPException(404, f"Uploaded file not found: {req.file_id!r}")

    meta = _read_upload_meta(req.file_id)
    job_id = store.create_job(
        job_type="data_gen",
        file_id=req.file_id,
        source_filename=meta.get("filename", excel_path.name),
        mode=req.data_mode,
    )
    store.add_version(
        job_id,
        kind="source",
        src_path=excel_path,
        filename=meta.get("filename", excel_path.name),
    )

    threading.Thread(target=_run_data_gen_job, args=(job_id, req), daemon=True).start()
    return {"job_id": job_id}


# ── Jobs & history ─────────────────────────────────────────────

@app.get("/api/jobs")
def list_jobs():
    return store.list_jobs()


@app.get("/api/history")
def list_history():
    items = store.list_jobs()
    summaries = []
    for j in items:
        summaries.append(
            {
                "job_id": j["job_id"],
                "job_type": j.get("job_type", "pipeline"),
                "status": j["status"],
                "mode": j.get("mode"),
                "source_filename": j.get("source_filename"),
                "download_filename": j.get("download_filename"),
                "created_at": j.get("created_at"),
                "completed_at": j.get("completed_at"),
                "total": (j.get("result") or {}).get("total"),
                "version_count": len(j.get("versions", [])),
            }
        )
    return summaries


@app.get("/api/jobs/{job_id}")
def get_job(job_id: str):
    job = store.get_job(job_id)
    if job is None:
        raise HTTPException(404, "Job not found")
    return job


@app.patch("/api/jobs/{job_id}")
def patch_job(job_id: str, body: PatchJobRequest):
    job = store.get_job(job_id)
    if job is None:
        raise HTTPException(404, "Job not found")
    if body.download_filename is not None:
        name = _safe_download_name(body.download_filename)
        store.update_job(job_id, download_filename=name)
    return store.get_job(job_id)


@app.get("/api/jobs/{job_id}/workbook")
def get_job_workbook(job_id: str, kind: str = "output", version_no: Optional[int] = None):
    job = store.get_job(job_id)
    if job is None:
        raise HTTPException(404, "Job not found")

    if version_no is not None:
        path = store.get_version_path(job_id, version_no)
    elif kind == "source":
        path = store.latest_version_path(job_id, "source") or store.latest_version_path(
            job_id, "edited_source"
        )
        if path is None:
            try:
                path = _find_upload(job["file_id"])
            except FileNotFoundError:
                raise HTTPException(404, "Source file not found")
    else:
        path = (
            store.latest_version_path(job_id, "edited_output")
            or store.latest_version_path(job_id, "output")
            or _resolve_output_path(job)
        )

    if path is None or not path.exists():
        raise HTTPException(404, "Workbook not found")

    return workbook_to_json(path)


@app.put("/api/jobs/{job_id}/workbook")
def save_job_workbook(job_id: str, payload: WorkbookPayload, kind: str = "output"):
    job = store.get_job(job_id)
    if job is None:
        raise HTTPException(404, "Job not found")

    version_kind = "edited_source" if kind == "source" else "edited_output"
    filename = payload.save_as_filename or (job.get("download_filename") or "edited.xlsx")
    filename = _safe_download_name(filename)

    job_dir = OUTPUT_DIR / job_id
    job_dir.mkdir(exist_ok=True)
    temp = job_dir / f"edit_{uuid.uuid4().hex[:8]}.xlsx"
    workbook_from_json(
        sheet_names=payload.sheet_names or [payload.active_sheet],
        active_sheet=payload.active_sheet,
        rows=payload.rows,
        dest=temp,
    )

    parent_no = None
    versions = job.get("versions", [])
    if versions:
        parent_no = max(v.get("version_no", 0) for v in versions)

    entry = store.add_version(
        job_id,
        kind=version_kind,
        src_path=temp,
        filename=filename,
        parent_version_no=parent_no,
    )

    if kind != "source":
        store.update_job(job_id, download_filename=filename)
        result = dict(job.get("result") or {})
        result["output_file"] = filename
        store.update_job(job_id, result=result)
        dest = job_dir / filename
        shutil.copy2(temp, dest)

    return {"version": entry, "download_filename": filename}


@app.get("/api/jobs/{job_id}/download")
def download_result(job_id: str, filename: Optional[str] = None, version_no: Optional[int] = None):
    job = store.get_job(job_id)
    if job is None:
        raise HTTPException(404, "Job not found")
    if job["status"] != "completed":
        raise HTTPException(400, f"Job not completed (status: {job['status']})")

    if version_no is not None:
        output_path = store.get_version_path(job_id, version_no)
    else:
        output_path = (
            store.latest_version_path(job_id, "edited_output")
            or store.latest_version_path(job_id, "output")
            or _resolve_output_path(job)
        )

    if output_path is None or not output_path.exists():
        raise HTTPException(500, "Output file missing on server")

    download_name = filename or job.get("download_filename") or output_path.name
    download_name = _safe_download_name(download_name)

    return FileResponse(
        output_path,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        filename=download_name,
    )


# ── TLS Client Bot intents ─────────────────────────────────────

class TlsIntentPayload(BaseModel):
    tag: str
    patterns: list[str] = Field(default_factory=list)
    responses: list[str] = Field(default_factory=list)


class TlsIntentUpdatePayload(BaseModel):
    tag: Optional[str] = None
    patterns: Optional[list[str]] = None
    responses: Optional[list[str]] = None


class TlsTrainRequest(BaseModel):
    epochs: int = 300


class PromptUpdateBody(BaseModel):
    content: str


@app.get("/api/tls-bot/prompts/{kind}")
def tls_bot_get_prompt(kind: str):
    if kind not in ("llm", "hybrid"):
        raise HTTPException(400, "kind must be 'llm' or 'hybrid'")
    return get_prompt(kind)  # type: ignore[arg-type]


@app.put("/api/tls-bot/prompts/{kind}")
def tls_bot_save_prompt(kind: str, body: PromptUpdateBody):
    if kind not in ("llm", "hybrid"):
        raise HTTPException(400, "kind must be 'llm' or 'hybrid'")
    if not body.content.strip():
        raise HTTPException(400, "Prompt content cannot be empty")
    try:
        return save_prompt(kind, body.content)  # type: ignore[arg-type]
    except KeyError as exc:
        raise HTTPException(400, str(exc))


@app.post("/api/tls-bot/prompts/{kind}/reset")
def tls_bot_reset_prompt(kind: str):
    if kind not in ("llm", "hybrid"):
        raise HTTPException(400, "kind must be 'llm' or 'hybrid'")
    return reset_prompt(kind)  # type: ignore[arg-type]


@app.get("/api/tls-bot/stats")
def tls_bot_stats():
    return tls_bot_service.get_stats()


@app.get("/api/tls-bot/intents")
def tls_bot_list_intents(q: str = ""):
    return {
        "intents": tls_bot_service.list_intents(q),
        "stats": tls_bot_service.get_stats(),
    }


@app.get("/api/tls-bot/intents/{tag}")
def tls_bot_get_intent(tag: str):
    try:
        return tls_bot_service.get_intent(tag)
    except KeyError:
        raise HTTPException(404, f"Intent not found: {tag!r}")


@app.post("/api/tls-bot/intents")
def tls_bot_create_intent(body: TlsIntentPayload):
    try:
        return tls_bot_service.create_intent(body.model_dump())
    except ValueError as exc:
        raise HTTPException(400, str(exc))


@app.put("/api/tls-bot/intents/{tag}")
def tls_bot_update_intent(tag: str, body: TlsIntentUpdatePayload):
    try:
        payload: dict[str, Any] = {}
        if body.tag is not None:
            payload["tag"] = body.tag
        if body.patterns is not None:
            payload["patterns"] = body.patterns
        if body.responses is not None:
            payload["responses"] = body.responses
        return tls_bot_service.update_intent(tag, payload)
    except KeyError:
        raise HTTPException(404, f"Intent not found: {tag!r}")
    except ValueError as exc:
        raise HTTPException(400, str(exc))


@app.delete("/api/tls-bot/intents/{tag}")
def tls_bot_delete_intent(tag: str):
    try:
        tls_bot_service.delete_intent(tag)
        return {"ok": True, "deleted": tag}
    except KeyError:
        raise HTTPException(404, f"Intent not found: {tag!r}")


# ── Query-param variants (safe for tags containing "/" or other special chars) ──
# The Next.js rewrite layer decodes %2F in path segments back to "/", breaking
# path-param routes for tags like "Không hỗ trợ / không thanh toán".
# Query strings are NOT re-parsed for path separators, so %2F stays intact.

@app.put("/api/tls-bot/intent")
def tls_bot_update_intent_q(
    tag: str = Query(..., description="Original tag of the intent to update"),
    body: TlsIntentUpdatePayload = ...,
):
    try:
        payload: dict[str, Any] = {}
        if body.tag is not None:
            payload["tag"] = body.tag
        if body.patterns is not None:
            payload["patterns"] = body.patterns
        if body.responses is not None:
            payload["responses"] = body.responses
        return tls_bot_service.update_intent(tag, payload)
    except KeyError:
        raise HTTPException(404, f"Intent not found: {tag!r}")
    except ValueError as exc:
        raise HTTPException(400, str(exc))


@app.delete("/api/tls-bot/intent")
def tls_bot_delete_intent_q(
    tag: str = Query(..., description="Tag of the intent to delete"),
):
    try:
        tls_bot_service.delete_intent(tag)
        return {"ok": True, "deleted": tag}
    except KeyError:
        raise HTTPException(404, f"Intent not found: {tag!r}")


@app.post("/api/tls-bot/retrain")
def tls_bot_retrain(body: TlsTrainRequest):
    try:
        return tls_bot_service.start_train(body.epochs)
    except RuntimeError as exc:
        raise HTTPException(409, str(exc))


class TlsSnapshotRequest(BaseModel):
    label: Optional[str] = None


@app.get("/api/tls-bot/versions")
def tls_bot_list_versions(limit: int = 50):
    return {"versions": tls_bot_service.list_versions(limit=limit)}


@app.post("/api/tls-bot/versions/snapshot")
def tls_bot_create_snapshot(body: TlsSnapshotRequest):
    entry = tls_bot_service.create_manual_snapshot(label=body.label)
    return entry


@app.get("/api/tls-bot/versions/{version_id}")
def tls_bot_get_version(version_id: str):
    try:
        return tls_bot_service.get_version(version_id)
    except KeyError:
        raise HTTPException(404, f"Version not found: {version_id!r}")
    except FileNotFoundError as exc:
        raise HTTPException(500, str(exc))


@app.post("/api/tls-bot/versions/{version_id}/restore")
def tls_bot_restore_version(version_id: str):
    try:
        return tls_bot_service.restore_version(version_id)
    except KeyError:
        raise HTTPException(404, f"Version not found: {version_id!r}")
    except FileNotFoundError as exc:
        raise HTTPException(500, str(exc))


@app.get("/api/tls-bot/train-status")
def tls_bot_train_status():
    return tls_bot_service.get_train_status()


@app.get("/api/jobs/{job_id}/versions/{version_no}/download")
def download_version(job_id: str, version_no: int):
    job = store.get_job(job_id)
    if job is None:
        raise HTTPException(404, "Job not found")
    path = store.get_version_path(job_id, version_no)
    if path is None:
        raise HTTPException(404, "Version not found")
    v = next((x for x in job.get("versions", []) if x.get("version_no") == version_no), None)
    name = v["filename"] if v else path.name
    return FileResponse(
        path,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        filename=name,
    )


# ── Static frontend (production) ──────────────────────────────────────────────
# The React app is built into /app/frontend/dist by the Dockerfile.
# All /api/* routes above take priority; everything else falls through to the SPA.
_DIST = Path(__file__).resolve().parent.parent / "frontend" / "dist"
if _DIST.is_dir():
    app.mount("/assets", StaticFiles(directory=_DIST / "assets"), name="assets")

    @app.api_route("/{full_path:path}", methods=["GET", "HEAD"], response_class=HTMLResponse, include_in_schema=False)
    def spa_fallback(full_path: str):
        return (_DIST / "index.html").read_text()
