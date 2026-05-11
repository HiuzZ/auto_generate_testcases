from __future__ import annotations

import argparse
import json
import re
from pathlib import Path
from typing import Any, Iterable

import openpyxl
from openpyxl.cell.rich_text import CellRichText, TextBlock
from openpyxl.cell.text import InlineFont
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.worksheet import Worksheet
from openpyxl.formatting.rule import Rule
from openpyxl.styles.differential import DifferentialStyle
from openpyxl.worksheet.datavalidation import DataValidation

# ---------------------------------------------------------------------------
# Constants for header styling
# ---------------------------------------------------------------------------
HEADER_FILL = PatternFill(start_color="0016365C", end_color="0016365C", fill_type="solid")
HEADER_FONT = Font(bold=True, color="00FFFFFF")

# Full list of columns A → M
HEADERS = [
    "TC_ID",
    "Conditions",
    "Test Scenario",
    "Bot Responses",
    "Path",
    "Expected Action Code",
    "Test Data",
    "User Attributes",
    "Call Id",
    "Test Results",
    "Error Type",
    "Error Description",
    "FPT Comment",
]

# Section header style (for grouping rows)
SECTION_FONT = Font(bold=True, color="000000")
SECTION_FILL = PatternFill(fill_type="solid", fgColor="ffff00")

# Group fill colors – only used when use_group_fills=True
GROUP_FILLS = [
    PatternFill(fill_type="solid", fgColor="00FDE2E2"),
    PatternFill(fill_type="solid", fgColor="00E3F2FD"),
    PatternFill(fill_type="solid", fgColor="00E8F5E9"),
    PatternFill(fill_type="solid", fgColor="00E3F2FD"),
]

RED_INLINE_FONT = InlineFont(color="00FF0000")

# Border style
THIN_BORDER = Border(
    left=Side(style="thin"),
    right=Side(style="thin"),
    top=Side(style="thin"),
    bottom=Side(style="thin"),
)

# Column widths
COL_WIDTHS = {
    1: 10,   # TC_ID
    2: 40,   # Conditions
    3: 40,   # Test Scenario
    4: 90,   # Bot Responses
    5: 25,   # Path
    6: 25,   # Expected Action Code
    7: 40,   # Test Data
    8: 20,   # User Attributes
    9: 20,   # Call Id
    10: 20,   # Test Results
    11: 20,   # Error Type
    12: 20,   # Error Description
    13: 20,   # FPT Comment
}

# Conditional formatting colors for Test Results (column 10)
RESULT_COLORS = {
    "Pass": "00C6EFCE",        # light green
    "Fail": "00FFC7CE",        # light red
    "Pending": "00FFEB9C",     # light yellow
    "Fixed": "00FFEB9C",
    "Un-executed": "00FFEB9C",
    "Need review": "00D8BFD8", # light purple
    "Todo": "00FFFFFF",        # white
}
ERROR_TYPE_COLORS = {
    "ASR": "00FFFFFF",
    "Intent": "00FFFFFF",
    "Flow": "00FFFFFF",
    "Script": "00FFFFFF",
    "Other": "00FFFFFF",
}


def _load_testcases(path: Path) -> list[dict[str, Any]]:
    """Read a JSON array of test cases and return a list of cleaned dicts."""
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
        bot_responses = item.get("bot_responses", [])
        expected_action_code = item.get("expected_action_code", "N/A")
        tc_path = str(item.get("path", ""))
        highlight_last_step = bool(item.get("highlight_last_step", False))
        test_data = str(item.get("test_data", ""))

        if not isinstance(steps, Iterable):
            raise ValueError(f"'steps' for {tc_id} must be a list.")
        if not isinstance(bot_responses, Iterable):
            raise ValueError(f"'bot_responses' for {tc_id} must be a list.")

        # Remove cycle markers like "(cycle to A1)"
        clean_steps: list[str] = []
        for s in steps:
            s_str = str(s)
            if s_str.strip().startswith("(cycle to"):
                continue
            clean_steps.append(s_str)

        clean_bot_responses: list[str] = [str(r) for r in bot_responses]
        while len(clean_bot_responses) < len(clean_steps):
            clean_bot_responses.append("")

        cases.append({
            "tc_id": tc_id,
            "conditions": conditions,
            "steps": clean_steps,
            "bot_responses": clean_bot_responses,
            "expected_action_code": expected_action_code,
            "path": tc_path,
            "highlight_last_step": highlight_last_step,
            "test_data": test_data,
        })
    return cases


def _auto_fit_column(ws: Worksheet, col_idx: int, extra: int = 2) -> None:
    col_letter = get_column_letter(col_idx)
    max_len = 0
    for cell in ws[col_letter]:
        if cell.value:
            for part in str(cell.value).split("\n"):
                max_len = max(max_len, len(part))
    ws.column_dimensions[col_letter].width = max_len + extra


def _auto_fit_row_heights(ws: Worksheet, min_height: float = 15.0, line_height: float = 15.0) -> None:
    """Adjust row heights based on content line count."""
    for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
        row_idx = row[0].row
        max_lines = 1
        for cell in row:
            if cell.value:
                lines = str(cell.value).split("\n")
                max_lines = max(max_lines, len(lines))
        ws.row_dimensions[row_idx].height = max(max_lines * line_height, min_height)


_STEP_NO_RE = re.compile(r"\bA(\d+)\b", re.IGNORECASE)


def _extract_group_step(path_text: str) -> str | None:
    """Extract the grouping step from a path string.
    - If path contains "End", return the A<N> token immediately before "End".
    - Otherwise return the last A<N> token.
    """
    if not path_text:
        return None
    tokens = re.split(r"\s*(?:->|→|,)\s*", path_text.strip())
    tokens = [t.strip() for t in tokens if t.strip()]
    end_idx = next((i for i, t in enumerate(tokens) if t.lower() == "end"), -1)
    if end_idx != -1:
        if end_idx > 0:
            candidate = tokens[end_idx - 1]
            if _STEP_NO_RE.fullmatch(candidate):
                return candidate
        return None
    else:
        last = tokens[-1] if tokens else ""
        return last if _STEP_NO_RE.fullmatch(last) else None


def _step_sort_key(step_no: str) -> tuple[int, str]:
    m = _STEP_NO_RE.fullmatch(step_no.strip())
    return (int(m.group(1)), step_no) if m else (10**9, step_no)


def _load_step_name_map(testcases_path: Path, rows_path: Path | None = None) -> dict[str, str]:
    """Try to load step_no → step_name mapping from a rows JSON file."""
    candidates: list[Path] = []
    if rows_path is not None and rows_path.exists():
        candidates.append(rows_path)
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
                if step_no and step_name and step_name.upper() != step_no.upper() and step_no not in mapping:
                    mapping[step_no] = step_name
            if mapping:
                return mapping
        except Exception:
            continue
    return {}


def _apply_borders(ws: Worksheet) -> None:
    """Apply thin borders to all used cells."""
    for row in ws.iter_rows(min_row=1, max_row=ws.max_row, max_col=ws.max_column):
        for cell in row:
            cell.border = THIN_BORDER


def _set_column_widths(ws: Worksheet) -> None:
    """Set column widths according to COL_WIDTHS."""
    for col_idx, width in COL_WIDTHS.items():
        ws.column_dimensions[get_column_letter(col_idx)].width = width


def _apply_alignments(ws: Worksheet, section_header_rows: list[int] | None = None) -> None:
    """Vertical center for all cells, horizontal center for column A, J, K and row 1.
    Section header rows (if provided) are set to left alignment afterwards."""
    for row in ws.iter_rows(min_row=1, max_row=ws.max_row, max_col=ws.max_column):
        for cell in row:
            if cell.column == 1 or cell.column == 10 or cell.column == 11 or cell.row == 1:
                cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            else:
                cell.alignment = Alignment(vertical="center", wrap_text=True)

    # Override section header rows to left alignment
    if section_header_rows:
        for r in section_header_rows:
            ws.cell(row=r, column=1).alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)


def _add_data_validations(ws: Worksheet, max_row: int) -> None:
    """Add dropdown lists to columns J (Test Results) and K (Error Type)."""
    # Test Results list
    dv_results = DataValidation(
        type="list",
        formula1='"Todo,Pass,Fail,Pending,Fixed,Un-executed,Need review"',
        allow_blank=True,
    )
    dv_results.error = "Please select a value from the list."
    dv_results.errorTitle = "Invalid Result"
    ws.add_data_validation(dv_results)
    dv_results.add(f"J2:J{max_row}")

    # Error Type list
    dv_error = DataValidation(
        type="list",
        formula1='"ASR,Intent,Flow,Script,Other"',
        allow_blank=True,
    )
    dv_error.error = "Please select a value from the list."
    dv_error.errorTitle = "Invalid Error Type"
    ws.add_data_validation(dv_error)
    dv_error.add(f"K2:K{max_row}")


def _add_conditional_formatting(ws: Worksheet, max_row: int) -> None:
    """Apply background color rules to Test Results (J) and Error Type (K) columns."""
    # Test Results column
    for value, color in RESULT_COLORS.items():
        fill = PatternFill(start_color=color, end_color=color, fill_type="solid")
        dxf = DifferentialStyle(fill=fill)
        rule = Rule(
            type="cellIs",
            operator="equal",
            formula=[f'"{value}"'],
            dxf=dxf,
        )
        ws.conditional_formatting.add(f"J2:J{max_row}", rule)

    # Error Type column (white background for all choices, optional)
    for value, color in ERROR_TYPE_COLORS.items():
        fill = PatternFill(start_color=color, end_color=color, fill_type="solid")
        dxf = DifferentialStyle(fill=fill)
        rule = Rule(
            type="cellIs",
            operator="equal",
            formula=[f'"{value}"'],
            dxf=dxf,
        )
        ws.conditional_formatting.add(f"K2:K{max_row}", rule)


def export_to_excel(
    cases: list[dict[str, Any]],
    out_path: Path,
    step_name_map: dict[str, str] | None = None,
    *,
    group_by_step: bool = False,
    use_group_fills: bool = False,
    allow_highlight_last: bool = False,
) -> None:
    """Write test cases to an Excel file with full formatting."""
    step_name_map = step_name_map or {}
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "TestCases"

    # ---------- Write header row ----------
    for col_idx, header_text in enumerate(HEADERS, start=1):
        cell = ws.cell(row=1, column=col_idx, value=header_text)
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.alignment = Alignment(wrap_text=True, vertical="center")

    # ---------- Common helper for writing a data row ----------
    def _write_data_row(
        row: int, case: dict[str, Any], display_tc_index: int, fill: PatternFill | None = None
    ) -> int:
        """Write a single data row (columns A–M) and return the next row index."""
        tc_id = f"TC{display_tc_index:03d}"
        conditions = str(case.get("conditions", ""))
        steps: list[str] = case.get("steps", [])
        bot_responses: list[str] = case.get("bot_responses", [])
        expected_action = str(case.get("expected_action_code", "N/A"))
        tc_path = str(case.get("path", ""))
        highlight_last = bool(case.get("highlight_last_step", False))
        test_data = str(case.get("test_data", ""))

        numbered_steps = [f"{i}. {s}" for i, s in enumerate(steps, 1)]
        scenario = "\n".join(numbered_steps)

        numbered_responses: list[str] = []
        for i, r in enumerate(bot_responses, 1):
            numbered_responses.append(f"{i}. {r}" if r else f"{i}. ")
        bot_responses_text = "\n".join(numbered_responses)

        # Column A – TC_ID
        ws.cell(row=row, column=1, value=tc_id)
        # Column B – Conditions
        ws.cell(row=row, column=2, value=conditions)

        # Column C – Test Scenario
        cell_scn = ws.cell(row=row, column=3)
        if allow_highlight_last and highlight_last and numbered_steps:
            rt = CellRichText()
            if len(numbered_steps) > 1:
                rt.append("\n".join(numbered_steps[:-1]) + "\n")
            rt.append(TextBlock(RED_INLINE_FONT, numbered_steps[-1]))
            cell_scn.value = rt
        else:
            cell_scn.value = scenario

        # Column D – Bot Responses
        cell_bot = ws.cell(row=row, column=4)
        if allow_highlight_last and highlight_last and numbered_responses:
            rt_bot = CellRichText()
            if len(numbered_responses) > 1:
                rt_bot.append("\n".join(numbered_responses[:-1]) + "\n")
            rt_bot.append(TextBlock(RED_INLINE_FONT, numbered_responses[-1]))
            cell_bot.value = rt_bot
        else:
            cell_bot.value = bot_responses_text

        # Column E – Path
        ws.cell(row=row, column=5, value=tc_path)
        # Column F – Expected Action Code
        ws.cell(row=row, column=6, value=expected_action)
        # Column G – Test Data
        ws.cell(row=row, column=7, value=test_data)
        # Columns H–M remain empty (User Attributes, Call Id, Test Results, ...)

        if fill is not None:
            for c in range(1, 8):
                ws.cell(row=row, column=c).fill = fill

        return row + 1

    # ---------- Write test cases ----------
    section_header_rows: list[int] = []  # Keep track of group header rows

    if not group_by_step:
        row = 2
        for idx, case in enumerate(cases, start=1):
            row = _write_data_row(row, case, idx, fill=None)
    else:
        grouped_cases: dict[str, list[dict[str, Any]]] = {}
        for case in cases:
            group_step = _extract_group_step(str(case.get("path", ""))) or "UNKNOWN"
            grouped_cases.setdefault(group_step, []).append(case)

        ordered_groups = sorted(grouped_cases.keys(), key=_step_sort_key)
        current_group_key: tuple[Any, ...] | None = None
        group_index = -1
        display_tc_index = 1
        row = 2

        for step_no in ordered_groups:
            step_name = step_name_map.get(step_no, step_no)
            section_label = f"{step_no} - {step_name}"

            # Insert section header row (merge across all columns)
            ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=13)
            cell_sec = ws.cell(row=row, column=1, value=section_label)
            cell_sec.font = SECTION_FONT
            cell_sec.fill = SECTION_FILL
            section_header_rows.append(row)   # remember this row index
            row += 1

            for case in grouped_cases[step_no]:
                conditions = str(case.get("conditions", ""))
                steps: list[str] = case.get("steps", [])
                tc_path = str(case.get("path", ""))
                expected_action = str(case.get("expected_action_code", "N/A"))

                if use_group_fills:
                    group_key = (conditions, tuple(steps), tc_path, expected_action)
                    if group_key != current_group_key:
                        group_index += 1
                        current_group_key = group_key
                    fill = GROUP_FILLS[group_index % len(GROUP_FILLS)]
                else:
                    fill = None

                row = _write_data_row(row, case, display_tc_index, fill=fill)
                display_tc_index += 1

    # ---------- Apply formatting ----------
    _apply_borders(ws)
    _set_column_widths(ws)
    _apply_alignments(ws, section_header_rows=section_header_rows)   # section headers become left-aligned
    _auto_fit_row_heights(ws)
    _add_data_validations(ws, ws.max_row)
    _add_conditional_formatting(ws, ws.max_row)

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
    parser.add_argument(
        "--group-by-step",
        action="store_true",
        default=False,
        help="Group test cases by step and add section headers.",
    )
    parser.add_argument(
        "--use-group-fills",
        action="store_true",
        default=False,
        help="Apply alternating background colours to groups (requires --group-by-step).",
    )
    parser.add_argument(
        "--allow-highlight-last",
        action="store_true",
        default=False,
        help="Highlight last step in red (used by multi_responses pipeline).",
    )
    args = parser.parse_args()

    cases = _load_testcases(args.in_path)
    step_name_map = _load_step_name_map(args.in_path, rows_path=args.rows_path)
    export_to_excel(
        cases,
        args.out_path,
        step_name_map=step_name_map,
        group_by_step=args.group_by_step,
        use_group_fills=args.use_group_fills,
        allow_highlight_last=args.allow_highlight_last,
    )
    print(f"Wrote {len(cases)} test cases to Excel: {args.out_path}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())