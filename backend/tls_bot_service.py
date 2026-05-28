from __future__ import annotations

import json
import shutil
import threading
import uuid
from datetime import datetime
from pathlib import Path
from typing import Any, Callable

PROJECT_ROOT = Path(__file__).resolve().parent.parent
TLS_BOT_DIR = PROJECT_ROOT / "tls_client_bot"
INTENTS_PATH = TLS_BOT_DIR / "intents.json"
INTENTS_BACKUP_DIR = TLS_BOT_DIR / "intent_backups"
MANIFEST_PATH = INTENTS_BACKUP_DIR / "manifest.json"
MAX_VERSIONS = 80

_train_lock = threading.Lock()
_train_state: dict[str, Any] = {
    "status": "idle",
    "progress": "",
    "started_at": None,
    "completed_at": None,
    "error": None,
    "result": None,
}


def _load_raw() -> dict[str, Any]:
    if not INTENTS_PATH.exists():
        return {"intents": []}
    return json.loads(INTENTS_PATH.read_text(encoding="utf-8"))


def _stats_from_data(data: dict[str, Any]) -> dict[str, int]:
    intents = data.get("intents", [])
    return {
        "intent_count": len(intents),
        "pattern_count": sum(len(i.get("patterns", [])) for i in intents),
        "response_count": sum(len(i.get("responses", [])) for i in intents),
    }


def _load_manifest() -> dict[str, Any]:
    INTENTS_BACKUP_DIR.mkdir(parents=True, exist_ok=True)
    if MANIFEST_PATH.exists():
        return json.loads(MANIFEST_PATH.read_text(encoding="utf-8"))
    return {"versions": []}


def _save_manifest(manifest: dict[str, Any]) -> None:
    MANIFEST_PATH.write_text(
        json.dumps(manifest, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )


def _migrate_legacy_backups(manifest: dict[str, Any]) -> dict[str, Any]:
    """Import old intents_YYYYMMDD_HHMMSS.json files into manifest once."""
    if manifest.get("versions"):
        return manifest
    legacy = sorted(INTENTS_BACKUP_DIR.glob("intents_*.json"))
    versions: list[dict[str, Any]] = []
    for path in legacy:
        try:
            data = json.loads(path.read_text(encoding="utf-8"))
            stamp = path.stem.replace("intents_", "")
            vid = f"legacy_{stamp}"
            new_name = f"v_{stamp}.json"
            dest = INTENTS_BACKUP_DIR / new_name
            if not dest.exists():
                shutil.copy2(path, dest)
            versions.append(
                {
                    "id": vid,
                    "created_at": datetime.strptime(stamp, "%Y%m%d_%H%M%S").isoformat()
                    if len(stamp) >= 15
                    else datetime.utcnow().isoformat(),
                    "reason": "legacy_import",
                    "label": "Imported backup",
                    "filename": new_name,
                    **_stats_from_data(data),
                }
            )
        except (json.JSONDecodeError, OSError, ValueError):
            continue
    versions.sort(key=lambda v: v.get("created_at", ""), reverse=True)
    manifest["versions"] = versions
    if versions:
        _save_manifest(manifest)
    return manifest


def _prune_versions(manifest: dict[str, Any]) -> None:
    versions = manifest.get("versions", [])
    if len(versions) <= MAX_VERSIONS:
        return
    to_remove = versions[MAX_VERSIONS:]
    manifest["versions"] = versions[:MAX_VERSIONS]
    for entry in to_remove:
        path = INTENTS_BACKUP_DIR / entry.get("filename", "")
        if path.is_file():
            path.unlink(missing_ok=True)


def _create_version_snapshot(
    *,
    reason: str,
    label: str | None = None,
    source_path: Path | None = None,
) -> dict[str, Any]:
    """Save a point-in-time copy of intents (current file or explicit path)."""
    INTENTS_BACKUP_DIR.mkdir(parents=True, exist_ok=True)
    src = source_path or INTENTS_PATH
    if not src.exists():
        raise FileNotFoundError("No intents.json to snapshot")

    data = json.loads(src.read_text(encoding="utf-8"))
    stamp = datetime.utcnow().strftime("%Y%m%d_%H%M%S")
    short_id = stamp
    version_id = f"{stamp}_{_uuid_suffix()}"
    filename = f"v_{version_id}.json"
    dest = INTENTS_BACKUP_DIR / filename
    shutil.copy2(src, dest)

    entry = {
        "id": version_id,
        "created_at": datetime.utcnow().isoformat(),
        "reason": reason,
        "label": label or _reason_label(reason),
        "filename": filename,
        **_stats_from_data(data),
    }

    manifest = _migrate_legacy_backups(_load_manifest())
    manifest.setdefault("versions", [])
    manifest["versions"].insert(0, entry)
    _prune_versions(manifest)
    _save_manifest(manifest)
    return entry


def _uuid_suffix() -> str:
    return uuid.uuid4().hex[:8]


def _reason_label(reason: str) -> str:
    labels = {
        "manual_snapshot": "Manual backup",
        "before_create": "Before create intent",
        "before_update": "Before update intent",
        "before_delete": "Before delete intent",
        "before_restore": "Before rollback",
        "restored": "After rollback",
    }
    return labels.get(reason, reason.replace("_", " ").title())


def _save_raw(data: dict[str, Any], *, reason: str = "auto_save") -> None:
    """Persist intents; always snapshots the previous file before overwriting."""
    if INTENTS_PATH.exists():
        _create_version_snapshot(reason=reason)
    INTENTS_PATH.write_text(
        json.dumps(data, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )


def list_versions(limit: int = 50) -> list[dict[str, Any]]:
    manifest = _migrate_legacy_backups(_load_manifest())
    versions = manifest.get("versions", [])[:limit]
    current_stats = _stats_from_data(_load_raw())
    return [
        {
            **v,
            "is_live_match": (
                v.get("intent_count") == current_stats["intent_count"]
                and v.get("pattern_count") == current_stats["pattern_count"]
            ),
        }
        for v in versions
    ]


def create_manual_snapshot(label: str | None = None) -> dict[str, Any]:
    return _create_version_snapshot(
        reason="manual_snapshot",
        label=label or "Manual backup",
    )


def get_version(version_id: str) -> dict[str, Any]:
    manifest = _migrate_legacy_backups(_load_manifest())
    entry = next((v for v in manifest.get("versions", []) if v["id"] == version_id), None)
    if entry is None:
        raise KeyError(f"Version not found: {version_id!r}")
    path = INTENTS_BACKUP_DIR / entry["filename"]
    if not path.exists():
        raise FileNotFoundError(f"Backup file missing: {entry['filename']}")
    data = json.loads(path.read_text(encoding="utf-8"))
    return {"version": entry, "data": data}


def restore_version(version_id: str) -> dict[str, Any]:
    manifest = _migrate_legacy_backups(_load_manifest())
    entry = next((v for v in manifest.get("versions", []) if v["id"] == version_id), None)
    if entry is None:
        raise KeyError(f"Version not found: {version_id!r}")
    path = INTENTS_BACKUP_DIR / entry["filename"]
    if not path.exists():
        raise FileNotFoundError(f"Backup file missing: {entry['filename']}")

    if INTENTS_PATH.exists():
        _create_version_snapshot(
            reason="before_restore",
            label=f"Before rollback to {entry.get('label', version_id)}",
        )

    shutil.copy2(path, INTENTS_PATH)
    restored_stats = _stats_from_data(_load_raw())
    return {
        "restored_from": entry,
        "current_stats": restored_stats,
        "message": "intents.json restored. Retrain the TLS bot if you use the trained model.",
    }


def _normalize_tag_spacing(tag: str) -> str:
    """Normalise whitespace around '/' separators.

    Ensures 'A / B', 'A /B', 'A/ B', and 'A/B' all become 'A / B',
    so intents that differ only in slash-spacing are treated as identical.
    """
    parts = [p.strip() for p in tag.split("/")]
    return " / ".join(parts)


def _normalize_intent(raw: dict[str, Any]) -> dict[str, Any]:
    tag = _normalize_tag_spacing(str(raw.get("tag", "")).strip())
    patterns = [str(p).strip() for p in raw.get("patterns", []) if str(p).strip()]
    responses = [str(r).strip() for r in raw.get("responses", []) if str(r).strip()]
    if not tag:
        raise ValueError("Tag is required")
    if not patterns:
        raise ValueError("At least one pattern is required")
    if not responses:
        raise ValueError("At least one response is required")
    return {"tag": tag, "patterns": patterns, "responses": responses}


def _find_index(intents: list[dict[str, Any]], tag: str) -> int:
    norm_tag = _normalize_tag_spacing(tag)
    for i, item in enumerate(intents):
        if _normalize_tag_spacing(item["tag"]) == norm_tag:
            return i
    raise KeyError(f"Intent not found: {tag!r}")


def _matches_search(intent: dict[str, Any], q: str) -> bool:
    if not q:
        return True
    needle = q.lower()
    if needle in intent["tag"].lower():
        return True
    for p in intent["patterns"]:
        if needle in p.lower():
            return True
    for r in intent["responses"]:
        if needle in r.lower():
            return True
    return False


def get_stats() -> dict[str, int]:
    data = _load_raw()
    intents = data.get("intents", [])
    return {
        "intent_count": len(intents),
        "pattern_count": sum(len(i.get("patterns", [])) for i in intents),
        "response_count": sum(len(i.get("responses", [])) for i in intents),
    }


def list_intents(search: str = "") -> list[dict[str, Any]]:
    data = _load_raw()
    intents = data.get("intents", [])

    # Deduplicate by normalised tag — covers exact duplicates AND tags that
    # differ only in whitespace around '/' (e.g. "A / B" vs "A /B").
    # The last occurrence wins so the most-recently-edited version is kept.
    seen: dict[str, dict[str, Any]] = {}
    for intent in intents:
        norm_key = _normalize_tag_spacing(intent.get("tag", ""))
        seen[norm_key] = intent
    unique_intents = list(seen.values())

    q = search.strip()
    filtered = [i for i in unique_intents if _matches_search(i, q)]
    return sorted(filtered, key=lambda x: x.get("tag", "").lower())


def get_intent(tag: str) -> dict[str, Any]:
    data = _load_raw()
    idx = _find_index(data["intents"], tag)
    return dict(data["intents"][idx])


def create_intent(payload: dict[str, Any]) -> dict[str, Any]:
    intent = _normalize_intent(payload)
    data = _load_raw()
    norm_new = _normalize_tag_spacing(intent["tag"])
    if any(_normalize_tag_spacing(i["tag"]) == norm_new for i in data["intents"]):
        raise ValueError(f"Intent tag already exists: {intent['tag']!r}")
    data["intents"].append(intent)
    _save_raw(data, reason="before_create")
    return intent


def update_intent(tag: str, payload: dict[str, Any]) -> dict[str, Any]:
    data = _load_raw()
    idx = _find_index(data["intents"], tag)
    current = dict(data["intents"][idx])

    new_tag = str(payload.get("tag", current["tag"])).strip()
    patterns = payload.get("patterns", current["patterns"])
    responses = payload.get("responses", current["responses"])

    updated = _normalize_intent(
        {"tag": new_tag, "patterns": patterns, "responses": responses}
    )

    norm_new_tag = _normalize_tag_spacing(new_tag)
    norm_old_tag = _normalize_tag_spacing(tag)
    if norm_new_tag != norm_old_tag:
        if any(
            _normalize_tag_spacing(i["tag"]) == norm_new_tag
            for i in data["intents"]
            if _normalize_tag_spacing(i["tag"]) != norm_old_tag
        ):
            raise ValueError(f"Intent tag already exists: {new_tag!r}")

    data["intents"][idx] = updated
    _save_raw(data, reason="before_update")
    return updated


def delete_intent(tag: str) -> None:
    data = _load_raw()
    idx = _find_index(data["intents"], tag)
    data["intents"].pop(idx)
    _save_raw(data, reason="before_delete")


def get_train_status() -> dict[str, Any]:
    with _train_lock:
        return dict(_train_state)


def _set_train(**kwargs: Any) -> None:
    with _train_lock:
        _train_state.update(kwargs)


def _ensure_training_deps() -> None:
    """Verify NLTK/torch are installed and download tokenizer data if needed."""
    try:
        import nltk  # noqa: F401
    except ImportError as exc:
        raise ImportError(
            "Missing Python package 'nltk'. Install TLS bot dependencies:\n"
            "  cd auto_generate_test_cases && pip install -r requirements.txt"
        ) from exc

    try:
        import torch  # noqa: F401
    except ImportError as exc:
        raise ImportError(
            "Missing Python package 'torch'. Install TLS bot dependencies:\n"
            "  cd auto_generate_test_cases && pip install -r requirements.txt"
        ) from exc

    import nltk

    for resource in ("punkt_tab", "punkt"):
        try:
            nltk.data.find(f"tokenizers/{resource}")
            break
        except LookupError:
            continue
    else:
        _set_train(progress="Downloading NLTK tokenizer data…")
        try:
            nltk.download("punkt_tab", quiet=True)
        except Exception:
            nltk.download("punkt", quiet=True)


def _run_train(epochs: int) -> None:
    try:
        _set_train(
            status="running",
            progress="Preparing training data…",
            error=None,
            result=None,
        )

        _ensure_training_deps()

        import sys

        bot_dir = str(TLS_BOT_DIR)
        if bot_dir not in sys.path:
            sys.path.append(bot_dir)  # append, not insert(0) — avoids shadowing main excel_to_json

        from train import train_bot  # noqa: WPS433 — runtime import in bot dir

        logs: list[str] = []

        def log_fn(msg: str) -> None:
            logs.append(msg)
            _set_train(progress=msg)

        result = train_bot(
            epochs=epochs,
            intents_path=INTENTS_PATH,
            output_path=TLS_BOT_DIR / "data.pth",
            log=log_fn,
        )
        result["log"] = logs[-20:]

        _set_train(
            status="completed",
            progress="Training complete",
            completed_at=datetime.utcnow().isoformat(),
            result=result,
        )
    except Exception as exc:
        _set_train(
            status="failed",
            progress="",
            completed_at=datetime.utcnow().isoformat(),
            error=str(exc),
        )


def start_train(epochs: int = 300) -> dict[str, Any]:
    epochs = max(50, min(2000, epochs))
    with _train_lock:
        if _train_state["status"] == "running":
            raise RuntimeError("Training already in progress")

    _set_train(
        status="pending",
        progress="Starting…",
        started_at=datetime.utcnow().isoformat(),
        completed_at=None,
        error=None,
        result=None,
    )

    threading.Thread(target=_run_train, args=(epochs,), daemon=True).start()
    return get_train_status()
