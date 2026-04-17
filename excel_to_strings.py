from __future__ import annotations

import argparse
import json
from pathlib import Path

import pandas as pd


CANONICAL_COLUMNS = [
    "Step no",
    "Step name",
    "Customer intent",
    "Bot response",
    "Next Step",
    "Action code",
]

COLUMN_ALIASES: dict[str, list[str]] = {
    "Step no": ["step no", "step number", "no", "step"],
    "Step name": ["step name", "name", "topic"],
    "Customer intent": ["customer intent", "intent"],
    "Bot response": ["bot response", "response"],
    "Next Step": ["next step", "nest step", "next action"],
    "Action code": ["action code", "action_code"],
}

JSON_KEYS: dict[str, str] = {
    "Step no": "step_no",
    "Step name": "step_name",
    "Customer intent": "customer_intent",
    "Bot response": "bot_response",
    "Next Step": "next_step",
    "Action code": "action_code",
}


def _normalize_col(col: str) -> str:
    return " ".join(str(col).strip().split())


def _normalize_col_key(col: str) -> str:
    return _normalize_col(col).lower()


def _resolve_column_mapping(actual_columns: list[str]) -> dict[str, str]:
    norm_to_actual: dict[str, str] = {}
    for c in actual_columns:
        norm_to_actual[_normalize_col_key(c)] = c

    resolved: dict[str, str] = {}
    for canonical in CANONICAL_COLUMNS:
        aliases = COLUMN_ALIASES[canonical]
        found_actual: str | None = None

        for a in aliases:
            if a in norm_to_actual:
                found_actual = norm_to_actual[a]
                break

        if found_actual is None and canonical == "Action code":
            for norm, actual in norm_to_actual.items():
                if norm.startswith("action code"):
                    found_actual = actual
                    break

        if found_actual is None:
            raise ValueError(
                f"Missing required column for '{canonical}'. Present columns: {actual_columns}"
            )

        resolved[canonical] = found_actual

    return resolved


def _pick_excel_file(input_dir: Path) -> Path:
    candidates = sorted(
        [*input_dir.glob("*.xlsx"), *input_dir.glob("*.xlsm"), *input_dir.glob("*.xls")]
    )
    if not candidates:
        raise FileNotFoundError(
            f"No Excel file found in {input_dir}. Expected one of: *.xlsx, *.xlsm, *.xls"
        )
    return candidates[0]


def _detect_sheet_and_header_row(excel_path: Path) -> tuple[str | int, int]:
    xls = pd.ExcelFile(excel_path)
    required_alias_sets = []
    for canonical in CANONICAL_COLUMNS:
        aliases = set(COLUMN_ALIASES[canonical])
        if canonical == "Action code":
            aliases.add("action code")
        required_alias_sets.append((canonical, aliases))

    for sheet_name in xls.sheet_names:
        preview = pd.read_excel(
            excel_path,
            sheet_name=sheet_name,
            header=None,
            nrows=50,
            dtype=object,
        )
        for row_idx in range(len(preview)):
            row_values = [_normalize_col_key(v) for v in preview.iloc[row_idx].tolist()]
            row_set = set(row_values)

            ok = True
            for canonical, aliases in required_alias_sets:
                if canonical == "Action code":
                    if not any(v.startswith("action code") or v in aliases for v in row_set):
                        ok = False
                        break
                else:
                    if row_set.isdisjoint(aliases):
                        ok = False
                        break

            if ok:
                return sheet_name, row_idx

    raise ValueError(
        "Could not find a sheet/header row that matches the expected template columns. "
        f"Expected canonical fields: {CANONICAL_COLUMNS}"
    )


def _clean_dataframe(df: pd.DataFrame, mapping: dict[str, str]) -> pd.DataFrame:
    """Shared cleaning logic for both string and JSON converters."""
    df = df[[mapping[c] for c in CANONICAL_COLUMNS]].copy()
    df = df.rename(columns={v: k for k, v in mapping.items()})

    # Default for Action code is N/A.
    df["Action code"] = df["Action code"].where(df["Action code"].notna(), "N/A")
    df["Action code"] = df["Action code"].astype(str).str.strip()
    df.loc[df["Action code"].eq("") | df["Action code"].eq("nan"), "Action code"] = "N/A"

    # Convert all other fields to clean strings (empty if NaN).
    for col in ["Step no", "Step name", "Customer intent", "Bot response", "Next Step"]:
        df[col] = df[col].where(df[col].notna(), "")
        df[col] = df[col].astype(str).str.strip()
        df.loc[df[col].eq("nan"), col] = ""

    return df


def _load_dataframe(excel_path: Path, sheet: str | int | None) -> pd.DataFrame:
    """Load and rename columns from Excel, auto-detecting sheet/header if needed."""
    if sheet is None:
        detected_sheet, header_row = _detect_sheet_and_header_row(excel_path)
        df = pd.read_excel(
            excel_path, sheet_name=detected_sheet, header=header_row, dtype=object
        )
    else:
        df = pd.read_excel(excel_path, sheet_name=sheet, dtype=object)

    df = df.rename(columns={c: _normalize_col(c) for c in df.columns})
    return df


def convert_excel_rows_to_strings(
    excel_path: Path, sheet: str | int | None = None
) -> list[str]:
    """
    Read the Excel template and convert each row into:
    "Step no: <...>; Step name: <...>; Customer intent: <...>; Bot response: <...>; Next Step: <...>; Action code: <...>"
    """
    df = _load_dataframe(excel_path, sheet)
    mapping = _resolve_column_mapping(list(df.columns))
    df = _clean_dataframe(df, mapping)

    output: list[str] = []
    for _, row in df.iterrows():
        formatted = (
            f"Step no: {row['Step no']}; "
            f"Step name: {row['Step name']}; "
            f"Customer intent: {row['Customer intent']}; "
            f"Bot response: {row['Bot response']}; "
            f"Next Step: {row['Next Step']}; "
            f"Action code: {row['Action code']}"
        )
        output.append(formatted)

    return output


def convert_excel_rows_to_json(
    excel_path: Path, sheet: str | int | None = None
) -> list[dict[str, str]]:
    """
    Convert each row into a JSON-friendly object:
    {
      "step_no": "...",
      "step_name": "...",
      "customer_intent": "...",
      "bot_response": "...",
      "next_step": "...",
      "action_code": "..."
    }
    """
    df = _load_dataframe(excel_path, sheet)
    mapping = _resolve_column_mapping(list(df.columns))
    df = _clean_dataframe(df, mapping)

    rows: list[dict[str, str]] = []
    for _, r in df.iterrows():
        obj = {
            JSON_KEYS["Step no"]: r["Step no"],
            JSON_KEYS["Step name"]: r["Step name"],
            JSON_KEYS["Customer intent"]: r["Customer intent"],
            JSON_KEYS["Bot response"]: r["Bot response"],
            JSON_KEYS["Next Step"]: r["Next Step"],
            JSON_KEYS["Action code"]: r["Action code"],
        }
        rows.append(obj)

    return rows


def main() -> int:
    here = Path(__file__).resolve().parent
    default_input_dir = here / "input"

    parser = argparse.ArgumentParser(description="Convert Excel template rows into JSON objects.")
    parser.add_argument(
        "--file",
        type=Path,
        default=None,
        help="Path to an Excel file. If omitted, uses the first Excel file in ./input.",
    )
    parser.add_argument(
        "--sheet",
        default=None,
        help="Sheet name or index. If omitted, auto-detects the template sheet/header row.",
    )
    parser.add_argument(
        "--out",
        type=Path,
        default=None,
        help="Optional output JSON path. If provided, writes the full list to this file.",
    )
    args = parser.parse_args()

    excel_path = args.file if args.file is not None else _pick_excel_file(default_input_dir)
    sheet = (
        int(args.sheet)
        if isinstance(args.sheet, str) and args.sheet.isdigit()
        else args.sheet
    )

    rows = convert_excel_rows_to_json(excel_path, sheet=sheet)

    if args.out is not None:
        args.out.parent.mkdir(parents=True, exist_ok=True)
        args.out.write_text(json.dumps(rows, ensure_ascii=False, indent=2), encoding="utf-8")

    print(f"Loaded {len(rows)} rows from: {excel_path}")
    for obj in rows[:3]:
        print(json.dumps(obj, ensure_ascii=False))

    return 0


if __name__ == "__main__":
    raise SystemExit(main())