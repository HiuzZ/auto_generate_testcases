from __future__ import annotations

import argparse
import json
from pathlib import Path
from typing import Any, Iterable

import openpyxl
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.worksheet import Worksheet


def _load_testcases(path: Path) -> list[dict[str, Any]]:
    data = json.loads(path.read_text(encoding="utf-8"))
    if not isinstance(data, list):
        raise ValueError("Input JSON must be a list of test cases.")
    cases: list[dict[str, Any]] = []
    for i, item in enumerate(data):
        if not isinstance(item, dict):
            raise ValueError(f"Item at index {i} is not an object.")
        tc_id = str(item.get("tc_id", f"TC{i+1}"))
        steps = item.get("steps", [])
        bot_responses = item.get("bot_responses", [])   # new field
        expected_action_code = item.get("expected_action_code", "N/A")
        tc_path = str(item.get("path", ""))
        if not isinstance(steps, Iterable):
            raise ValueError(f"'steps' for {tc_id} must be a list.")
        if not isinstance(bot_responses, Iterable):
            raise ValueError(f"'bot_responses' for {tc_id} must be a list.")

        # Filter out cycle markers like "(cycle to A1)".
        clean_steps: list[str] = []
        for s in steps:
            s_str = str(s)
            if s_str.strip().startswith("(cycle to"):
                continue
            clean_steps.append(s_str)

        # Keep bot responses as they are (may be empty strings)
        clean_bot_responses: list[str] = [str(r) for r in bot_responses]

        # Ensure bot_responses length matches steps length (pad with empty strings)
        while len(clean_bot_responses) < len(clean_steps):
            clean_bot_responses.append("")

        cases.append(
            {
                "tc_id": tc_id,
                "steps": clean_steps,
                "bot_responses": clean_bot_responses,
                "expected_action_code": expected_action_code,
                "path": tc_path,
            }
        )
    return cases


def _auto_fit_column(ws: Worksheet, col_idx: int, extra: int = 2) -> None:
    col_letter = get_column_letter(col_idx)
    max_len = 0
    for cell in ws[col_letter]:
        if cell.value:
            for part in str(cell.value).split("\n"):
                max_len = max(max_len, len(part))
    ws.column_dimensions[col_letter].width = max_len + extra


def export_to_excel(cases: list[dict[str, Any]], out_path: Path) -> None:
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "TestCases"

    # Header
    ws["A1"] = "TC_ID"
    ws["B1"] = "Test Scenario"
    ws["C1"] = "Bot Responses"
    ws["D1"] = "Path"
    ws["E1"] = "Expected Action Code"
    ws["F1"] = "Test Data"

    row = 2
    for case in cases:
        tc_id = str(case["tc_id"])
        steps: list[str] = case.get("steps", [])
        bot_responses: list[str] = case.get("bot_responses", [])
        expected_action_code = str(case.get("expected_action_code", "N/A"))
        tc_path = str(case.get("path", ""))

        # Format steps as numbered list
        numbered_steps = [f"{idx}. {text}" for idx, text in enumerate(steps, start=1)]
        scenario = "\n".join(numbered_steps)

        # Format bot responses as numbered list (same numbering as steps)
        numbered_responses = []
        for idx, resp in enumerate(bot_responses, start=1):
            if resp:
                numbered_responses.append(f"{idx}. {resp}")
            else:
                numbered_responses.append(f"{idx}. ")
        bot_responses_text = "\n".join(numbered_responses)

        ws.cell(row=row, column=1, value=tc_id)
        cell_scn = ws.cell(row=row, column=2, value=scenario)
        cell_scn.alignment = openpyxl.styles.Alignment(wrap_text=True)

        cell_bot = ws.cell(row=row, column=3, value=bot_responses_text)
        cell_bot.alignment = openpyxl.styles.Alignment(wrap_text=True)

        ws.cell(row=row, column=4, value=tc_path)
        ws.cell(row=row, column=5, value=expected_action_code)
        # Column F ("Test Data") left empty for generate_test_data.py to fill later.

        row += 1

    # Auto-fit columns
    _auto_fit_column(ws, 1, extra=2)
    _auto_fit_column(ws, 2, extra=4)
    _auto_fit_column(ws, 3, extra=4)
    _auto_fit_column(ws, 4, extra=2)
    _auto_fit_column(ws, 5, extra=2)
    _auto_fit_column(ws, 6, extra=4)

    out_path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(out_path)


def main() -> int:
    parser = argparse.ArgumentParser(
        description="Convert JSON test cases into an Excel file (test_cases.xlsx)."
    )
    parser.add_argument(
        "--in",
        dest="in_path",
        type=Path,
        required=True,
        help="Input JSON file containing a list of test cases.",
    )
    parser.add_argument(
        "--out",
        dest="out_path",
        type=Path,
        default=Path("output/test_cases.xlsx"),
        help="Output .xlsx path (default: output/test_cases.xlsx).",
    )
    args = parser.parse_args()

    cases = _load_testcases(args.in_path)
    export_to_excel(cases, args.out_path)
    print(f"Wrote {len(cases)} test cases to Excel: {args.out_path}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())