from __future__ import annotations

import sys
import uuid
import threading
from pathlib import Path
from datetime import datetime
from typing import Any, Literal, Optional

# Import existing pipeline modules from parent directory
sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from fastapi import FastAPI, File, HTTPException, UploadFile
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse
from pydantic import BaseModel

UPLOAD_DIR = Path(__file__).parent / "uploads"
OUTPUT_DIR = Path(__file__).parent / "outputs"
UPLOAD_DIR.mkdir(parents=True, exist_ok=True)
OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

app = FastAPI(title="TC Generator API", version="1.0.0")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_methods=["*"],
    allow_headers=["*"],
)

_jobs: dict[str, dict[str, Any]] = {}
_jobs_lock = threading.Lock()


def _update_job(job_id: str, **kwargs: Any) -> None:
    with _jobs_lock:
        _jobs[job_id].update(kwargs)


# ──────────────────────────────────────────────────────────────
# Models
# ──────────────────────────────────────────────────────────────

class RunRequest(BaseModel):
    file_id: str
    mode: Literal["e2e_max", "e2e_short", "output_short", "multi_responses", "all"] = "all"
    sheet: Optional[str] = None
    root: Optional[str] = None
    max_depth: int = 200
    gen_data: bool = False


# ──────────────────────────────────────────────────────────────
# Health probes
# ──────────────────────────────────────────────────────────────

@app.get("/health/live")
def liveness():
    return {"status": "ok"}


@app.get("/health/ready")
def readiness():
    return {"status": "ok", "jobs_in_memory": len(_jobs)}


# ──────────────────────────────────────────────────────────────
# File upload
# ──────────────────────────────────────────────────────────────

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

    return {"file_id": file_id, "filename": file.filename, "size": dest.stat().st_size}


# ──────────────────────────────────────────────────────────────
# Pipeline background worker
# ──────────────────────────────────────────────────────────────

MODES_ALL = [
    ("e2e_short",       "Test case"),
    ("output_short",    "Test Output"),
    ("multi_responses", "Test case Đa Thoại"),
]
MODE_TO_SHEET = {m: s for m, s in MODES_ALL}


def _find_upload(file_id: str) -> Path:
    for ext in (".xlsx", ".xlsm", ".xls"):
        p = UPLOAD_DIR / f"{file_id}{ext}"
        if p.exists():
            return p
    raise FileNotFoundError(f"Uploaded file not found: {file_id!r}")


def _run_job(job_id: str, req: RunRequest) -> None:
    try:
        _update_job(job_id, status="running", progress="Locating uploaded file…")
        excel_path = _find_upload(req.file_id)

        from pipeline_tc import run_pipeline
        from excel_to_json import read_data_schema_sheet
        from tc_to_excel import _add_checklist_sheet, _add_data_input_sheet, _add_data_schema_sheet
        import openpyxl as xl

        job_dir = OUTPUT_DIR / job_id
        job_dir.mkdir(exist_ok=True)

        sheet = int(req.sheet) if req.sheet and req.sheet.isdigit() else req.sheet

        if req.mode == "all":
            schema_keys, schema_value_rows, schema_all_rows = read_data_schema_sheet(excel_path)
            combined_wb = xl.Workbook()
            combined_wb.remove(combined_wb.active)
            _add_checklist_sheet(combined_wb)

            results = []
            for mode, tc_sheet_name in MODES_ALL:
                _update_job(job_id, progress=f"Running {mode}…")
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

            _update_job(job_id, progress="Saving combined workbook…")
            _add_data_input_sheet(
                combined_wb,
                schema_keys=schema_keys or None,
                schema_value_rows=schema_value_rows or None,
            )
            _add_data_schema_sheet(combined_wb, schema_all_rows=schema_all_rows or None)

            output_excel = job_dir / "testcases_all.xlsx"
            combined_wb.save(output_excel)

            _update_job(
                job_id,
                status="completed",
                progress="Done",
                completed_at=datetime.utcnow().isoformat(),
                result={
                    "modes": results,
                    "total": sum(r["count"] for r in results),
                    "output_file": "testcases_all.xlsx",
                },
            )
        else:
            tc_sheet_name = MODE_TO_SHEET[req.mode]
            _update_job(job_id, progress=f"Running {req.mode}…")
            _, _, _, count, _ = run_pipeline(
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
            _update_job(
                job_id,
                status="completed",
                progress="Done",
                completed_at=datetime.utcnow().isoformat(),
                result={
                    "modes": [{"mode": req.mode, "count": count, "sheet": tc_sheet_name}],
                    "total": count,
                    "output_file": f"testcases_{req.mode}.xlsx",
                },
            )

    except Exception as exc:
        _update_job(
            job_id,
            status="failed",
            progress="",
            completed_at=datetime.utcnow().isoformat(),
            error=str(exc),
        )


# ──────────────────────────────────────────────────────────────
# Pipeline endpoints
# ──────────────────────────────────────────────────────────────

@app.post("/api/run")
def start_pipeline(req: RunRequest):
    try:
        _find_upload(req.file_id)
    except FileNotFoundError:
        raise HTTPException(404, f"Uploaded file not found: {req.file_id!r}")

    job_id = str(uuid.uuid4())
    with _jobs_lock:
        _jobs[job_id] = {
            "job_id": job_id,
            "status": "pending",
            "progress": "",
            "created_at": datetime.utcnow().isoformat(),
            "completed_at": None,
            "error": None,
            "result": None,
        }

    threading.Thread(target=_run_job, args=(job_id, req), daemon=True).start()
    return {"job_id": job_id}


@app.get("/api/jobs")
def list_jobs():
    with _jobs_lock:
        return sorted(_jobs.values(), key=lambda j: j["created_at"], reverse=True)


@app.get("/api/jobs/{job_id}")
def get_job(job_id: str):
    with _jobs_lock:
        job = _jobs.get(job_id)
    if job is None:
        raise HTTPException(404, "Job not found")
    return job


@app.get("/api/jobs/{job_id}/download")
def download_result(job_id: str):
    with _jobs_lock:
        job = _jobs.get(job_id)
    if job is None:
        raise HTTPException(404, "Job not found")
    if job["status"] != "completed":
        raise HTTPException(400, f"Job not completed (status: {job['status']})")

    output_path = OUTPUT_DIR / job_id / job["result"]["output_file"]
    if not output_path.exists():
        raise HTTPException(500, "Output file missing on server")

    return FileResponse(
        output_path,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        filename=job["result"]["output_file"],
    )
