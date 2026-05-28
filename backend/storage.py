from __future__ import annotations

import json
import shutil
import threading
import uuid
from datetime import datetime
from pathlib import Path
from typing import Any, Literal

VersionKind = Literal["source", "output", "edited_source", "edited_output"]


class JobStore:
    def __init__(self, data_dir: Path, versions_dir: Path) -> None:
        self.data_dir = data_dir
        self.versions_dir = versions_dir
        self.jobs_file = data_dir / "jobs.json"
        self._lock = threading.Lock()
        self._jobs: dict[str, dict[str, Any]] = {}
        self.data_dir.mkdir(parents=True, exist_ok=True)
        self.versions_dir.mkdir(parents=True, exist_ok=True)
        self._load()

    def _load(self) -> None:
        if not self.jobs_file.exists():
            return
        try:
            raw = json.loads(self.jobs_file.read_text(encoding="utf-8"))
            if isinstance(raw, dict):
                self._jobs = raw
        except (json.JSONDecodeError, OSError):
            self._jobs = {}

    def _persist(self) -> None:
        self.jobs_file.write_text(
            json.dumps(self._jobs, ensure_ascii=False, indent=2),
            encoding="utf-8",
        )

    def create_job(
        self,
        *,
        job_type: Literal["pipeline", "data_gen"],
        file_id: str,
        source_filename: str,
        mode: str,
    ) -> str:
        job_id = str(uuid.uuid4())
        job = {
            "job_id": job_id,
            "job_type": job_type,
            "status": "pending",
            "progress": "",
            "progress_pct": 0,
            "created_at": datetime.utcnow().isoformat(),
            "completed_at": None,
            "error": None,
            "file_id": file_id,
            "source_filename": source_filename,
            "mode": mode,
            "download_filename": None,
            "result": None,
            "versions": [],
        }
        with self._lock:
            self._jobs[job_id] = job
            self._persist()
        return job_id

    def update_job(self, job_id: str, **kwargs: Any) -> None:
        with self._lock:
            if job_id not in self._jobs:
                return
            self._jobs[job_id].update(kwargs)
            self._persist()

    def get_job(self, job_id: str) -> dict[str, Any] | None:
        with self._lock:
            job = self._jobs.get(job_id)
            return dict(job) if job else None

    def list_jobs(self, limit: int = 100) -> list[dict[str, Any]]:
        with self._lock:
            items = sorted(
                self._jobs.values(),
                key=lambda j: j.get("created_at", ""),
                reverse=True,
            )
        return [dict(j) for j in items[:limit]]

    def add_version(
        self,
        job_id: str,
        *,
        kind: VersionKind,
        src_path: Path,
        filename: str,
        parent_version_no: int | None = None,
    ) -> dict[str, Any]:
        job_dir = self.versions_dir / job_id
        job_dir.mkdir(parents=True, exist_ok=True)
        with self._lock:
            job = self._jobs.get(job_id)
            if not job:
                raise KeyError(job_id)
            version_no = len(job.get("versions", [])) + 1
            safe_name = filename.replace("/", "_").replace("\\", "_")
            dest = job_dir / f"v{version_no}_{kind}_{safe_name}"
            shutil.copy2(src_path, dest)
            entry = {
                "version_id": str(uuid.uuid4()),
                "version_no": version_no,
                "kind": kind,
                "filename": filename,
                "stored_name": dest.name,
                "created_at": datetime.utcnow().isoformat(),
                "parent_version_no": parent_version_no,
            }
            job.setdefault("versions", []).append(entry)
            self._persist()
            return dict(entry)

    def get_version_path(self, job_id: str, version_no: int) -> Path | None:
        job = self.get_job(job_id)
        if not job:
            return None
        for v in job.get("versions", []):
            if v.get("version_no") == version_no:
                p = self.versions_dir / job_id / v["stored_name"]
                if p.exists():
                    return p
        return None

    def latest_version_path(self, job_id: str, kind: VersionKind | None = None) -> Path | None:
        job = self.get_job(job_id)
        if not job:
            return None
        versions = job.get("versions", [])
        if kind:
            versions = [v for v in versions if v.get("kind") == kind]
        if not versions:
            return None
        latest = max(versions, key=lambda v: v.get("version_no", 0))
        p = self.versions_dir / job_id / latest["stored_name"]
        return p if p.exists() else None
