from __future__ import annotations

import argparse
import json
import re
from pathlib import Path
from typing import Any, Iterable

import openpyxl
from openpyxl.cell.rich_text import CellRichText, TextBlock
from openpyxl.cell.text import InlineFont
from openpyxl.styles import Font, PatternFill
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
        conditions = str(item.get("conditions", ""))
        steps = item.get("steps", [])
        bot_responses = item.get("bot_responses", [])   # new field
        expected_action_code = item.get("expected_action_code", "N/A")
        tc_path = str(item.get("path", ""))
        highlight_last_step = bool(item.get("highlight_last_step", False))
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
                "conditions": conditions,
                "steps": clean_steps,
                "bot_responses": clean_bot_responses,
                "expected_action_code": expected_action_code,
                "path": tc_path,
                "highlight_last_step": highlight_last_step,
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


_STEP_NO_RE = re.compile(r"\bA(\d+)\b", re.IGNORECASE)


def _extract_max_step_no(path_text: str) -> str | None:
    nums = [int(m.group(1)) for m in _STEP_NO_RE.finditer(path_text or "")]
    if not nums:
        return None
    return f"A{max(nums)}"


def _step_sort_key(step_no: str) -> tuple[int, str]:
    m = _STEP_NO_RE.fullmatch(step_no.strip())
    if m:
        return (int(m.group(1)), step_no)
    return (10**9, step_no)


def _load_step_name_map(testcases_path: Path, rows_path: Path | None = None) -> dict[str, str]:
    if rows_path is not None and rows_path.exists():
        candidates: list[Path] = [rows_path]
    else:
        candidates = []

    # Try to infer matching rows file from testcase filename.
    in_name = testcases_path.name
    if in_name.startswith("testcases_"):
        candidates.append(testcases_path.with_name(in_name.replace("testcases_", "rows_", 1)))
    candidates.append(testcases_path.with_name("rows_multi_vb.json"))
    candidates.append(testcases_path.with_name("rows_e2e_vb.json"))

    for candidate in candidates:
        if not candidate.exists():
            continue
        try:
            data = json.loads(candidate.read_text(encoding="utf-8"))
            if not isinstance(data, list):
                continue
            mapping: dict[str, str] = {}
            for item in data:
                if not isinstance(item, dict):
                    continue
                step_no = str(item.get("step_no", "")).strip()
                step_name = str(item.get("step_name", "")).strip()
                # Ignore degenerate values like step_name == step_no ("A1").
                if step_no and step_name and step_name.upper() != step_no.upper() and step_no not in mapping:
                    mapping[step_no] = step_name
            if mapping:
                return mapping
        except Exception:
            continue
    return {}


def export_to_excel(cases: list[dict[str, Any]], out_path: Path, step_name_map: dict[str, str] | None = None) -> None:
    step_name_map = step_name_map or {}
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "TestCases"

    # Header
    ws["A1"] = "TC_ID"
    ws["B1"] = "Conditions"
    ws["C1"] = "Test Scenario"
    ws["D1"] = "Bot Responses"
    ws["E1"] = "Path"
    ws["F1"] = "Expected Action Code"
    ws["G1"] = "Test Data"

    row = 2
    red_font = InlineFont(color="00FF0000")
    section_font = Font(bold=True, color="001F4E78")
    section_fill = PatternFill(fill_type="solid", fgColor="00EEF3FB")
    group_fills = [
        PatternFill(fill_type="solid", fgColor="00FDE2E2"),  # light red
        PatternFill(fill_type="solid", fgColor="00E3F2FD"),  # light blue
        PatternFill(fill_type="solid", fgColor="00E8F5E9"),  # light green
        PatternFill(fill_type="solid", fgColor="00E3F2FD"),  # light blue
    ]

    grouped_cases: dict[str, list[dict[str, Any]]] = {}
    for case in cases:
        max_step = _extract_max_step_no(str(case.get("path", ""))) or "UNKNOWN"
        grouped_cases.setdefault(max_step, []).append(case)

    ordered_groups = sorted(grouped_cases.keys(), key=_step_sort_key)
    current_group_key: tuple[Any, ...] | None = None
    group_index = -1
    display_tc_index = 1

    for step_no in ordered_groups:
        step_name = step_name_map.get(step_no, step_no)
        section_label = f"{step_no} - {step_name}"
        ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=6)
        ws.cell(row=row, column=1, value=section_label)
        ws.cell(row=row, column=1).font = section_font
        ws.cell(row=row, column=1).fill = section_fill
        ws.cell(row=row, column=1).alignment = openpyxl.styles.Alignment(wrap_text=True)
        row += 1

        # Preserve the incoming case order inside each section.
        # This lets each pipeline control adjacency in its own output,
        # while Excel still groups by the section header step.
        section_cases = grouped_cases[step_no]

        for case in section_cases:
            tc_id = f"TC{display_tc_index:03d}"
            conditions = str(case.get("conditions", ""))
            steps: list[str] = case.get("steps", [])
            bot_responses: list[str] = case.get("bot_responses", [])
            expected_action_code = str(case.get("expected_action_code", "N/A"))
            tc_path = str(case.get("path", ""))
            highlight_last_step = bool(case.get("highlight_last_step", False))

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

            group_key = (conditions, tuple(steps), tc_path, expected_action_code)
            if group_key != current_group_key:
                group_index += 1
                current_group_key = group_key
            fill = group_fills[group_index % len(group_fills)]

            ws.cell(row=row, column=1, value=tc_id)
            ws.cell(row=row, column=2, value=conditions)
            cell_scn = ws.cell(row=row, column=3)
            if highlight_last_step and numbered_steps:
                rich_text = CellRichText()
                if len(numbered_steps) > 1:
                    rich_text.append("\n".join(numbered_steps[:-1]) + "\n")
                rich_text.append(TextBlock(red_font, numbered_steps[-1]))
                cell_scn.value = rich_text
            else:
                cell_scn.value = scenario
            cell_scn.alignment = openpyxl.styles.Alignment(wrap_text=True)

            cell_bot = ws.cell(row=row, column=4)
            if highlight_last_step and numbered_responses:
                rich_text_bot = CellRichText()
                if len(numbered_responses) > 1:
                    rich_text_bot.append("\n".join(numbered_responses[:-1]) + "\n")
                rich_text_bot.append(TextBlock(red_font, numbered_responses[-1]))
                cell_bot.value = rich_text_bot
            else:
                cell_bot.value = bot_responses_text
            cell_bot.alignment = openpyxl.styles.Alignment(wrap_text=True)

            ws.cell(row=row, column=5, value=tc_path)
            ws.cell(row=row, column=6, value=expected_action_code)
            # Column G ("Test Data") left empty for generate_test_data.py to fill later.
            for col in range(1, 7):
                ws.cell(row=row, column=col).fill = fill

            row += 1
            display_tc_index += 1

    # Auto-fit columns
    _auto_fit_column(ws, 1, extra=2)
    _auto_fit_column(ws, 2, extra=2)
    _auto_fit_column(ws, 3, extra=4)
    _auto_fit_column(ws, 4, extra=4)
    _auto_fit_column(ws, 5, extra=2)
    _auto_fit_column(ws, 6, extra=2)
    _auto_fit_column(ws, 7, extra=4)

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
    parser.add_argument(
        "--rows",
        dest="rows_path",
        type=Path,
        default=None,
        help="Optional rows JSON (with step_no/step_name) for exact section titles.",
    )
    args = parser.parse_args()

    cases = _load_testcases(args.in_path)
    step_name_map = _load_step_name_map(args.in_path, rows_path=args.rows_path)
    export_to_excel(cases, args.out_path, step_name_map=step_name_map)
    print(f"Wrote {len(cases)} test cases to Excel: {args.out_path}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
