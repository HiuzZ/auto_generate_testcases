from __future__ import annotations

import argparse
import json
import re
from pathlib import Path

import pandas as pd


CANONICAL_COLUMNS = [
    "Step no",
    "Step name",
    "Conditions",
    "Customer intent",
    "Next Step",
    "Bot response",
    "Bot response 2",
    "Bot response 3",
    "Bot response 4",
    "Bot response 5",
]

REQUIRED_CANONICAL_COLUMNS = [
    "Step no",
    "Step name",
    "Customer intent",
    "Bot response",
    "Next Step",
]

# Normalized (lowercased, collapsed whitespace) aliases for each canonical field.
# This covers common variants found in the workbook (e.g. "Nest Step" and "Next action").
COLUMN_ALIASES: dict[str, list[str]] = {
    "Step no": ["step no", "step number", "no", "step"],
    "Step name": ["step name", "name", "topic"],
    "Conditions": ["conditions", "condition"],
    "Customer intent": ["customer intent", "intent"],
    "Bot response": ["bot response", "response", "bot_response"],
    "Bot response 2": ["bot response 2", "response 2", "bot_response_2"],
    "Bot response 3": ["bot response 3", "response 3", "bot_response_3"],
    "Bot response 4": ["bot response 4", "response 4", "bot_response_4"],
    "Bot response 5": ["bot response 5", "response 5", "bot_response_5"],
    "Next Step": ["next step", "nest step", "next action"],
}

JSON_KEYS: dict[str, str] = {
    "Step no": "step_no",
    "Step name": "step_name",
    "Conditions": "conditions",
    "Customer intent": "customer_intent",
    "Bot response": "bot_response",
    "Bot response 2": "bot_response_2",
    "Bot response 3": "bot_response_3",
    "Bot response 4": "bot_response_4",
    "Bot response 5": "bot_response_5",
    "Next Step": "next_step",
}

FIXED_POSITIONAL_COLUMNS = list(CANONICAL_COLUMNS)

_DYNAMIC_KEY_RE = re.compile(r"\W+")


def _normalize_col(col: str) -> str:
    return " ".join(str(col).strip().split())


def _normalize_col_key(col: str) -> str:
    return _normalize_col(col).lower()


def _is_action_code_header(header: str) -> bool:
    key = _normalize_col_key(header)
    return key.startswith("action code") or key == "action_code"


def _json_key_from_header(header: str, existing: set[str]) -> str:
    is_action_code_column = _is_action_code_header(header)
    if is_action_code_column:
        base = "0"
    else:
        base = _DYNAMIC_KEY_RE.sub("_", _normalize_col(header).lower()).strip("_")
    if not base:
        base = "column"
    if base[0].isdigit() and not is_action_code_column:
        base = f"col_{base}"

    key = base
    suffix = 2
    while key in existing:
        key = f"{base}_{suffix}"
        suffix += 1
    existing.add(key)
    return key


def _resolve_column_mapping(actual_columns: list[str]) -> dict[str, str]:
    """
    Return a mapping {canonical: actual_column_name} based on COLUMN_ALIASES.
    Matching is done on normalized, lowercased column names.
    """
    norm_to_actual: dict[str, str] = {}
    for c in actual_columns:
        norm_to_actual[_normalize_col_key(c)] = c

    resolved: dict[str, str] = {}
    for canonical in REQUIRED_CANONICAL_COLUMNS:
        aliases = COLUMN_ALIASES[canonical]
        found_actual: str | None = None

        for a in aliases:
            if a in norm_to_actual:
                found_actual = norm_to_actual[a]
                break

        if found_actual is None:
            raise ValueError(
                f"Missing required column for '{canonical}'. Present columns: {actual_columns}"
            )

        resolved[canonical] = found_actual

    for optional_col in [
        "Conditions",
        "Bot response 2",
        "Bot response 3",
        "Bot response 4",
        "Bot response 5",
    ]:
        found_optional: str | None = None
        for alias in COLUMN_ALIASES[optional_col]:
            if alias in norm_to_actual:
                found_optional = norm_to_actual[alias]
                break
        if found_optional is not None:
            resolved[optional_col] = found_optional

    return resolved


def _is_fixed_positional_header(values: list[object]) -> bool:
    if len(values) < len(FIXED_POSITIONAL_COLUMNS):
        return False

    for value, canonical in zip(values[: len(FIXED_POSITIONAL_COLUMNS)], FIXED_POSITIONAL_COLUMNS):
        value_key = _normalize_col_key(value)
        aliases = COLUMN_ALIASES.get(canonical, [canonical.lower()])
        if canonical == "Next Step":
            if value_key not in aliases:
                return False
        elif value_key not in aliases:
            return False

    return True


def _detect_dynamic_columns(df: pd.DataFrame, start_idx: int = 10) -> dict[str, str]:
    dynamic_columns: dict[str, str] = {}
    existing = {*JSON_KEYS.values(), "action_code"}

    for col in list(df.columns)[start_idx:]:
        header = _normalize_col(col)
        if not header or header.lower().startswith("unnamed:"):
            continue

        series = df[col]
        has_data = series.notna() & series.astype(str).str.strip().ne("") & series.astype(str).str.lower().ne("nan")
        if not has_data.any():
            continue

        dynamic_columns[col] = _json_key_from_header(header, existing)

    return dynamic_columns


def _clean_rows_dataframe(df: pd.DataFrame, dynamic_columns: dict[str, str] | None = None) -> tuple[pd.DataFrame, dict[str, str]]:
    dynamic_columns = dynamic_columns or {}

    for col in CANONICAL_COLUMNS:
        df[col] = df[col].where(df[col].notna(), "")
        df[col] = df[col].astype(str).str.strip()
        df.loc[df[col].eq("nan"), col] = ""

    for col in dynamic_columns:
        df[col] = df[col].where(df[col].notna(), "")
        df[col] = df[col].astype(str).str.strip()
        df.loc[df[col].eq("nan"), col] = ""

    return df, dynamic_columns


def _read_template_dataframe(
    excel_path: Path,
    sheet: str | int | None = None,
) -> tuple[pd.DataFrame, dict[str, str]]:
    if sheet is None:
        detected_sheet, header_row = _detect_sheet_and_header_row(excel_path)
        df = pd.read_excel(
            excel_path, sheet_name=detected_sheet, header=header_row, dtype=object
        )
    else:
        df = pd.read_excel(excel_path, sheet_name=sheet, dtype=object)

    df = df.rename(columns={c: _normalize_col(c) for c in df.columns})

    if _is_fixed_positional_header(list(df.columns)):
        fixed_actual_columns = list(df.columns[: len(FIXED_POSITIONAL_COLUMNS)])
        rename_map = {
            actual: canonical
            for actual, canonical in zip(fixed_actual_columns, FIXED_POSITIONAL_COLUMNS)
        }
        dynamic_columns = _detect_dynamic_columns(df, start_idx=len(FIXED_POSITIONAL_COLUMNS))
        selected_columns = [*fixed_actual_columns, *dynamic_columns.keys()]
        df = df[selected_columns].copy()
        df = df.rename(columns=rename_map)
        return _clean_rows_dataframe(df, dynamic_columns)

    mapping = _resolve_column_mapping(list(df.columns))
    dynamic_columns = _detect_dynamic_columns(df, start_idx=len(FIXED_POSITIONAL_COLUMNS))

    selected_columns = [mapping[c] for c in REQUIRED_CANONICAL_COLUMNS]
    if "Conditions" in mapping:
        selected_columns.insert(2, mapping["Conditions"])
    insert_idx = selected_columns.index(mapping["Bot response"]) + 1
    for optional_response_col in ["Bot response 2", "Bot response 3", "Bot response 4", "Bot response 5"]:
        if optional_response_col in mapping:
            selected_columns.insert(insert_idx, mapping[optional_response_col])
        insert_idx += 1
    selected_columns.extend(col for col in dynamic_columns if col not in selected_columns)

    df = df[selected_columns].copy()
    df = df.rename(columns={v: k for k, v in mapping.items()})
    if "Conditions" not in df.columns:
        df["Conditions"] = ""
    for optional_response_col in ["Bot response 2", "Bot response 3", "Bot response 4", "Bot response 5"]:
        if optional_response_col not in df.columns:
            df[optional_response_col] = ""
    df = df[[*CANONICAL_COLUMNS, *dynamic_columns.keys()]]

    return _clean_rows_dataframe(df, dynamic_columns)


def _pick_excel_file(input_dir: Path) -> Path:
    candidates = sorted(
        [
            path
            for path in [*input_dir.glob("*.xlsx"), *input_dir.glob("*.xlsm"), *input_dir.glob("*.xls")]
            if not path.name.startswith("~$")
        ]
    )
    if not candidates:
        raise FileNotFoundError(
            f"No Excel file found in {input_dir}. Expected one of: *.xlsx, *.xlsm, *.xls"
        )
    return candidates[0]


def _detect_sheet_and_header_row(excel_path: Path) -> tuple[str | int, int]:
    """
    Many workbooks have multiple sheets and the template header might not start
    on the first row. This scans each sheet for a row containing all expected
    columns and returns (sheet_name, header_row_index).
    """
    xls = pd.ExcelFile(excel_path)
    # Prioritise the sheet named "Input" so two-sheet workbooks are handled correctly.
    _sheet_names = xls.sheet_names
    if "Input" in _sheet_names:
        _sheet_names = ["Input"] + [s for s in _sheet_names if s != "Input"]

    for sheet_name in _sheet_names:
        preview = pd.read_excel(
            excel_path,
            sheet_name=sheet_name,
            header=None,
            nrows=50,
            dtype=object,
        )
        for row_idx in range(len(preview)):
            if _is_fixed_positional_header(preview.iloc[row_idx].tolist()):
                return sheet_name, row_idx

    # We detect the header row by looking for the presence of at least one alias
    # for every required canonical field.
    required_alias_sets = []
    for canonical in REQUIRED_CANONICAL_COLUMNS:
        aliases = set(COLUMN_ALIASES[canonical])
        required_alias_sets.append((canonical, aliases))

    for sheet_name in _sheet_names:
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
                if row_set.isdisjoint(aliases):
                    ok = False
                    break

            if ok:
                return sheet_name, row_idx

    raise ValueError(
        "Could not find a sheet/header row that matches the expected template columns. "
        f"Expected canonical fields: {CANONICAL_COLUMNS}"
    )


def read_data_schema_sheet(
    excel_path: Path,
) -> tuple[list[str], list[list[str]], list[list[str]]]:
    """Read the 'Data Schema' sheet from an input Excel file.

    Layout expected (row 1 = header, row 2+ = data):
        A       B           C       D               E
        ID      Field       Type    Description     Note
        1       CALL_SCRIPT string  Loại kịch bản   ...
        2       CONTRACTID  string  Mã hợp đồng
        ...

    Returns (keys, value_rows, all_rows):
    - keys: non-empty values from column B rows 2+ (the Field names) — used as extra
            column headers in the Data input output sheet
    - value_rows: always [] — Data input rows are left empty for testers to fill in
    - all_rows: complete raw cell content for copying to the output Data Schema sheet
    """
    try:
        xls = pd.ExcelFile(str(excel_path))
    except Exception:
        return [], [], []

    schema_sheet: str | None = None
    for candidate in ("Data Schema", "Data schema"):
        if candidate in xls.sheet_names:
            schema_sheet = candidate
            break
    if schema_sheet is None:
        return [], [], []

    df = pd.read_excel(str(excel_path), sheet_name=schema_sheet, header=None, dtype=object)

    def _v(val: object) -> str:
        if val is None:
            return ""
        try:
            if pd.isna(val):  # type: ignore[arg-type]
                return ""
        except Exception:
            pass
        return str(val).strip()

    all_rows = [[_v(v) for v in df.iloc[row_idx].tolist()] for row_idx in range(len(df))]

    if len(df) < 2 or len(df.columns) < 2:
        return [], [], all_rows

    # Keys = non-empty values in column B (index 1), starting from row 2 (index 1, skipping header)
    keys: list[str] = [
        _v(df.iloc[row_idx, 1])
        for row_idx in range(1, len(df))
        if _v(df.iloc[row_idx, 1])
    ]
    return keys, [], all_rows


def convert_excel_rows_to_strings(
    excel_path: Path, sheet: str | int | None = None
) -> list[str]:
    """
    Legacy output (list of formatted strings). Prefer convert_excel_rows_to_json().
    """
    """
    Read the Excel template and convert each row into:
    "Step no: <...>; Step name: <...>; Customer intent: <...>; Bot response: <...>; Next Step: <...>; Action code: <...>"
    """
    df, _dynamic_columns = _read_template_dataframe(excel_path, sheet=sheet)

    output: list[str] = []
    for _, row in df.iterrows():
        action_code = " \\ ".join(
            str(row[col]).strip()
            for col in _dynamic_columns
            if str(row[col]).strip()
        ) or "N/A"
        formatted = (
            f"Step no: {row['Step no']}; "
            f"Step name: {row['Step name']}; "
            f"Conditions: {row['Conditions']}; "
            f"Customer intent: {row['Customer intent']}; "
            f"Bot response: {row['Bot response']}; "
            f"Bot response 2: {row['Bot response 2']}; "
            f"Bot response 3: {row['Bot response 3']}; "
            f"Bot response 4: {row['Bot response 4']}; "
            f"Bot response 5: {row['Bot response 5']}; "
            f"Next Step: {row['Next Step']}; "
            f"Action code: {action_code}"
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
      "action_code": "...",
      "...dynamic columns after J...": "..."
    }
    """
    df, dynamic_columns = _read_template_dataframe(excel_path, sheet=sheet)

    rows: list[dict[str, str]] = []
    for _, r in df.iterrows():
        action_values = [
            str(r[col]).strip()
            for col in dynamic_columns
            if str(r[col]).strip()
        ]
        obj = {
            JSON_KEYS["Step no"]: r["Step no"],
            JSON_KEYS["Step name"]: r["Step name"],
            JSON_KEYS["Conditions"]: r["Conditions"],
            JSON_KEYS["Customer intent"]: r["Customer intent"],
            JSON_KEYS["Bot response"]: r["Bot response"],
            JSON_KEYS["Bot response 2"]: r["Bot response 2"],
            JSON_KEYS["Bot response 3"]: r["Bot response 3"],
            JSON_KEYS["Bot response 4"]: r["Bot response 4"],
            JSON_KEYS["Bot response 5"]: r["Bot response 5"],
            JSON_KEYS["Next Step"]: r["Next Step"],
            "action_code": " \\ ".join(action_values) if action_values else "N/A",
        }
        for col, key in dynamic_columns.items():
            obj[key] = r[col]
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
