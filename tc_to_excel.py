from __future__ import annotations

import argparse
import json
import re
from pathlib import Path
from typing import Any, Iterable, List

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

# Main headers (fixed part, columns A–F only – column G becomes part of Data collection)
FIXED_MAIN_HEADERS = [
    "TC_ID \n (Mã TC)",
    "Conditions \n (Điều kiện kiểm tra)",
    "User Attributes \n (Bộ dữ liệu đầu vào)",
    "Test Scenario \n (Các bước thực hiện kiểm tra)",
    "Expected Bot Responses \n (Câu trả lời mong muốn của bot)",
    "Flow \n (Bước theo kịch bản)",
]

# Sub‑headers for each round (7 columns each)
ROUND_SUB_HEADERS = [
    "Tester\n(Người kiểm thử)",
    "Test Data \n (Dữ liệu kiểm tra)",
    "Test Results \n (Kết quả kiểm tra)",
    "Error Type \n (Loại lỗi)",
    "Error Description \n (Mô tả lỗi)",
    "Call ID",
    "FPT Comment \n (Ghi chú FPT)",
]

# Section header style (for grouping rows)
SECTION_FONT = Font(bold=True, color="000000")
SECTION_FILL = PatternFill(fill_type="solid", fgColor="ffff00")

# Group fill colors – only used when use_group_fills=True, applied to columns A-G
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

# Base column widths (indices relative to fixed columns)
BASE_COL_WIDTHS = {
    1: 10,   # TC_ID
    2: 40,   # Conditions
    3: 30,   # User Attributes
    4: 40,   # Test Scenario
    5: 90,   # Expected Bot Responses
    6: 25,   # Flow (Path)
}
# Column G (Expected Action Code) gets default width 25 (same as before)
# Default width for dynamic columns
DYN_COL_WIDTH = 20
# Round sub-column widths (same for both rounds)
ROUND_COL_WIDTHS = [15, 40, 20, 20, 20, 20, 20]  # Tester, Test Data, Test Results, Error Type, Error Description, Call ID, FPT Comment

INPUT_ROW_COLUMNS: list[tuple[str, str]] = [
    ("Step no", "step_no"),
    ("Step name", "step_name"),
    ("Conditions", "conditions"),
    ("Customer intent", "customer_intent"),
    ("Bot response", "bot_response"),
    ("Bot response 2", "bot_response_2"),
    ("Bot response 3", "bot_response_3"),
    ("Bot response 4", "bot_response_4"),
    ("Bot response 5", "bot_response_5"),
    ("Next Step", "next_step"),
    ("Action code", "action_code"),
]

INPUT_ROW_KEYS = {key for _, key in INPUT_ROW_COLUMNS}

# Conditional formatting colors for Test Results columns.
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
    "ASR Mistake": "00FFFFFF",
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
        test_data_round2 = str(item.get("test_data_round2", ""))
        source_columns = item.get("source_columns", {})

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
            "test_data_round2": test_data_round2,
            "source_columns": source_columns,
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


def _load_source_rows(rows_path: Path | None) -> list[dict[str, Any]] | None:
    if rows_path is None or not rows_path.exists():
        return None

    data = json.loads(rows_path.read_text(encoding="utf-8"))
    if not isinstance(data, list):
        raise ValueError(f"Rows JSON must be a list: {rows_path}")

    rows: list[dict[str, Any]] = []
    for idx, item in enumerate(data):
        if not isinstance(item, dict):
            raise ValueError(f"Rows JSON item at index {idx} is not an object: {rows_path}")
        rows.append(item)
    return rows


def _apply_borders(ws: Worksheet) -> None:
    """Apply thin borders to all used cells."""
    for row in ws.iter_rows(min_row=1, max_row=ws.max_row, max_col=ws.max_column):
        for cell in row:
            cell.border = THIN_BORDER


def _set_column_widths(ws: Worksheet, dynamic_keys: List[str]) -> None:
    """Set column widths based on BASE_COL_WIDTHS, dynamic keys, and round columns."""
    # Fixed columns 1-6
    for col_idx, width in BASE_COL_WIDTHS.items():
        ws.column_dimensions[get_column_letter(col_idx)].width = width
    # Column G (Expected Action Code) width
    ws.column_dimensions[get_column_letter(7)].width = 25
    # Dynamic columns (starting at column 8)
    for i, key in enumerate(dynamic_keys, start=8):
        ws.column_dimensions[get_column_letter(i)].width = DYN_COL_WIDTH
    # Round 1 columns
    round1_start = 8 + len(dynamic_keys)
    for i, width in enumerate(ROUND_COL_WIDTHS):
        ws.column_dimensions[get_column_letter(round1_start + i)].width = width
    # Round 2 columns
    round2_start = round1_start + len(ROUND_COL_WIDTHS)
    for i, width in enumerate(ROUND_COL_WIDTHS):
        ws.column_dimensions[get_column_letter(round2_start + i)].width = width


def _apply_alignments(ws: Worksheet, dynamic_keys: List[str], section_header_rows: list[int] | None = None) -> None:
    """Vertical center for all cells, horizontal center for column A, round results/error columns, and row 1.
    Section header rows (if provided) are set to left alignment afterwards."""
    # Calculate dynamic column indices
    round1_start = 8 + len(dynamic_keys)
    round1_result_col = round1_start + 2   # Test Results (index 2 within round)
    round1_error_col = round1_start + 3    # Error Type
    round2_start = round1_start + len(ROUND_COL_WIDTHS)
    round2_result_col = round2_start + 2
    round2_error_col = round2_start + 3

    for row in ws.iter_rows(min_row=1, max_row=ws.max_row, max_col=ws.max_column):
        for cell in row:
            if cell.column == 1 or cell.row == 2 or \
               cell.column == round1_result_col or cell.column == round1_error_col or \
               cell.column == round2_result_col or cell.column == round2_error_col or \
               cell.row == 1:
                cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            else:
                cell.alignment = Alignment(vertical="center", wrap_text=True)

    # Override section header rows to left alignment
    if section_header_rows:
        for r in section_header_rows:
            ws.cell(row=r, column=1).alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)


def _add_data_validations(ws: Worksheet, dynamic_keys: List[str], max_row: int) -> None:
    """Add dropdown lists to Test Results and Error Type columns for both rounds."""
    round1_start = 8 + len(dynamic_keys)
    round1_result_col = round1_start + 2
    round1_error_col = round1_start + 3
    round2_start = round1_start + len(ROUND_COL_WIDTHS)
    round2_result_col = round2_start + 2
    round2_error_col = round2_start + 3

    # Test Results list
    dv_results = DataValidation(
        type="list",
        formula1='"Todo,Pass,Fail,Pending,Fixed,Un-executed,Need review"',
        allow_blank=True,
    )
    dv_results.error = "Please select a value from the list."
    dv_results.errorTitle = "Invalid Result"
    ws.add_data_validation(dv_results)
    dv_results.add(f"{get_column_letter(round1_result_col)}2:{get_column_letter(round1_result_col)}{max_row}")
    dv_results.add(f"{get_column_letter(round2_result_col)}2:{get_column_letter(round2_result_col)}{max_row}")

    # Error Type list
    dv_error = DataValidation(
        type="list",
        formula1='"ASR Mistake,Intent,Flow,Script,Other"',
        allow_blank=True,
    )
    dv_error.error = "Please select a value from the list."
    dv_error.errorTitle = "Invalid Error Type"
    ws.add_data_validation(dv_error)
    dv_error.add(f"{get_column_letter(round1_error_col)}2:{get_column_letter(round1_error_col)}{max_row}")
    dv_error.add(f"{get_column_letter(round2_error_col)}2:{get_column_letter(round2_error_col)}{max_row}")


def _add_conditional_formatting(ws: Worksheet, dynamic_keys: List[str], max_row: int) -> None:
    """Apply background color rules to Test Results and Error Type columns for both rounds."""
    round1_start = 8 + len(dynamic_keys)
    round1_result_col = round1_start + 2
    round1_error_col = round1_start + 3
    round2_start = round1_start + len(ROUND_COL_WIDTHS)
    round2_result_col = round2_start + 2
    round2_error_col = round2_start + 3

    # Test Results columns
    for value, color in RESULT_COLORS.items():
        fill = PatternFill(start_color=color, end_color=color, fill_type="solid")
        dxf = DifferentialStyle(fill=fill)
        rule = Rule(
            type="cellIs",
            operator="equal",
            formula=[f'"{value}"'],
            dxf=dxf,
        )
        ws.conditional_formatting.add(f"{get_column_letter(round1_result_col)}2:{get_column_letter(round1_result_col)}{max_row}", rule)
        ws.conditional_formatting.add(f"{get_column_letter(round2_result_col)}2:{get_column_letter(round2_result_col)}{max_row}", rule)

    # Error Type columns
    for value, color in ERROR_TYPE_COLORS.items():
        fill = PatternFill(start_color=color, end_color=color, fill_type="solid")
        dxf = DifferentialStyle(fill=fill)
        rule = Rule(
            type="cellIs",
            operator="equal",
            formula=[f'"{value}"'],
            dxf=dxf,
        )
        ws.conditional_formatting.add(f"{get_column_letter(round1_error_col)}2:{get_column_letter(round1_error_col)}{max_row}", rule)
        ws.conditional_formatting.add(f"{get_column_letter(round2_error_col)}2:{get_column_letter(round2_error_col)}{max_row}", rule)


def _title_from_key(key: str) -> str:
    return " ".join(part for part in key.split("_") if part).title()


def _add_input_rows_sheet(wb: openpyxl.Workbook, source_rows: list[dict[str, Any]]) -> None:
    if not source_rows:
        return

    ws = wb.create_sheet("InputRows")
    dynamic_keys: list[str] = []
    for row in source_rows:
        for key in row:
            if key in INPUT_ROW_KEYS or key in dynamic_keys:
                continue
            dynamic_keys.append(key)

    headers = [header for header, _ in INPUT_ROW_COLUMNS] + [_title_from_key(key) for key in dynamic_keys]
    keys = [key for _, key in INPUT_ROW_COLUMNS] + dynamic_keys

    for col_idx, header in enumerate(headers, start=1):
        cell = ws.cell(row=1, column=col_idx, value=header)
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    for row_idx, row_data in enumerate(source_rows, start=2):
        for col_idx, key in enumerate(keys, start=1):
            ws.cell(row=row_idx, column=col_idx, value=str(row_data.get(key, "")))

    _apply_borders(ws)
    for col_idx, key in enumerate(keys, start=1):
        width = 18
        if key in {"conditions", "customer_intent", "bot_response", "bot_response_2", "bot_response_3", "bot_response_4", "bot_response_5"}:
            width = 45
        elif key in {"step_name", "action_code"}:
            width = 28
        ws.column_dimensions[get_column_letter(col_idx)].width = width

    for row in ws.iter_rows(min_row=1, max_row=ws.max_row, max_col=ws.max_column):
        for cell in row:
            cell.alignment = Alignment(vertical="center", wrap_text=True)

    ws.freeze_panes = "A2"


def export_to_excel(
    cases: list[dict[str, Any]],
    out_path: Path,
    step_name_map: dict[str, str] | None = None,
    *,
    group_by_step: bool = False,
    use_group_fills: bool = False,
    allow_highlight_last: bool = False,
    source_rows: list[dict[str, Any]] | None = None,
) -> None:
    """Write test cases to an Excel file with full formatting."""
    step_name_map = step_name_map or {}

    # Collect all dynamic keys from all test cases
    all_dynamic_keys = set()
    for case in cases:
        source_columns = case.get("source_columns", {})
        for k in source_columns:
            if k not in INPUT_ROW_KEYS:
                all_dynamic_keys.add(k)
    dynamic_keys = sorted(all_dynamic_keys)  # deterministic order

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "TestCases"

    # Calculate column ranges
    num_dynamic = len(dynamic_keys)
    round1_start = 8 + num_dynamic   # column after dynamic columns
    round2_start = round1_start + len(ROUND_COL_WIDTHS)
    total_columns = round2_start + len(ROUND_COL_WIDTHS) - 1

    # ---------- Write row 1 (top header row) ----------
    # 1. Fixed headers A-F
    for col_idx, header_text in enumerate(FIXED_MAIN_HEADERS, start=1):
        cell = ws.cell(row=1, column=col_idx, value=header_text)
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.alignment = Alignment(wrap_text=True, vertical="center")

    # 2. Merged "Data collection" group (columns G to last dynamic column)
    if num_dynamic > 0:
        data_collection_start = 7
        data_collection_end = 7 + num_dynamic
    else:
        # No dynamic columns: still create a single column group for G only
        data_collection_start = 7
        data_collection_end = 7
    ws.merge_cells(start_row=1, start_column=data_collection_start, end_row=1, end_column=data_collection_end)
    cell_data = ws.cell(row=1, column=data_collection_start, value="Data Collection")
    cell_data.fill = HEADER_FILL
    cell_data.font = HEADER_FONT
    cell_data.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    # 3. Round 1 merged header
    ws.merge_cells(start_row=1, start_column=round1_start, end_row=1, end_column=round1_start + len(ROUND_COL_WIDTHS) - 1)
    cell_r1 = ws.cell(row=1, column=round1_start, value="Round 1 (Vòng 1)")
    cell_r1.fill = HEADER_FILL
    cell_r1.font = HEADER_FONT
    cell_r1.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    # 4. Round 2 merged header
    ws.merge_cells(start_row=1, start_column=round2_start, end_row=1, end_column=round2_start + len(ROUND_COL_WIDTHS) - 1)
    cell_r2 = ws.cell(row=1, column=round2_start, value="Round 2 (Vòng 2)")
    cell_r2.fill = HEADER_FILL
    cell_r2.font = HEADER_FONT
    cell_r2.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    # ---------- Write row 2 (sub-header row) ----------
    # Fixed columns A-F are merged vertically, so we don't write anything there (they will be overwritten by the merge)
    # But we must ensure the vertical merge is applied. We'll do that after setting row2.
    # Column G: Expected Action Code
    ws.cell(row=2, column=7, value="action_code").fill = HEADER_FILL
    ws.cell(row=2, column=7).font = HEADER_FONT
    ws.cell(row=2, column=7).alignment = Alignment(wrap_text=True, vertical="center", horizontal="center")

    # Dynamic columns (starting at column 8) – each key as header
    for i, key in enumerate(dynamic_keys, start=8):
        cell = ws.cell(row=2, column=i, value=key)
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.alignment = Alignment(wrap_text=True, vertical="center", horizontal="center")

    # Round 1 sub-headers
    for i, sub_header in enumerate(ROUND_SUB_HEADERS):
        col_idx = round1_start + i
        cell = ws.cell(row=2, column=col_idx, value=sub_header)
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.alignment = Alignment(wrap_text=True, vertical="center", horizontal="center")

    # Round 2 sub-headers
    for i, sub_header in enumerate(ROUND_SUB_HEADERS):
        col_idx = round2_start + i
        cell = ws.cell(row=2, column=col_idx, value=sub_header)
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.alignment = Alignment(wrap_text=True, vertical="center", horizontal="center")

    # Merge fixed columns A–F vertically (row 1 and 2)
    for col_idx in range(1, 7):
        ws.merge_cells(start_row=1, start_column=col_idx, end_row=2, end_column=col_idx)

    # No vertical merge for column G or dynamic columns – they stay as individual cells.

    # ---------- Common helper for writing a data row ----------
    def _write_data_row(
        row: int, case: dict[str, Any], display_tc_index: int, fill: PatternFill | None = None
    ) -> int:
        """Write a single data row (columns A through total_columns) and return the next row index."""
        tc_id = f"TC{display_tc_index:03d}"
        conditions = str(case.get("conditions", ""))
        steps: list[str] = case.get("steps", [])
        bot_responses: list[str] = case.get("bot_responses", [])
        expected_action = str(case.get("expected_action_code", "N/A"))
        tc_path = str(case.get("path", ""))
        highlight_last = bool(case.get("highlight_last_step", False))
        test_data = str(case.get("test_data", ""))
        test_data_round2 = str(case.get("test_data_round2", ""))
        source_columns = case.get("source_columns", {})

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
        # Column C – User Attributes (kept empty as before)
        ws.cell(row=row, column=3, value="")
        # Column D – Test Scenario
        cell_scn = ws.cell(row=row, column=4)
        if allow_highlight_last and highlight_last and numbered_steps:
            rt = CellRichText()
            if len(numbered_steps) > 1:
                rt.append("\n".join(numbered_steps[:-1]) + "\n")
            rt.append(TextBlock(RED_INLINE_FONT, numbered_steps[-1]))
            cell_scn.value = rt
        else:
            cell_scn.value = scenario
        # Column E – Expected Bot Responses
        cell_bot = ws.cell(row=row, column=5)
        if allow_highlight_last and highlight_last and numbered_responses:
            rt_bot = CellRichText()
            if len(numbered_responses) > 1:
                rt_bot.append("\n".join(numbered_responses[:-1]) + "\n")
            rt_bot.append(TextBlock(RED_INLINE_FONT, numbered_responses[-1]))
            cell_bot.value = rt_bot
        else:
            cell_bot.value = bot_responses_text
        # Column F – Flow (Path)
        ws.cell(row=row, column=6, value=tc_path)
        # Column G – Expected Action Code
        ws.cell(row=row, column=7, value=expected_action)

        # Dynamic columns (after G)
        for i, key in enumerate(dynamic_keys):
            col_idx = 8 + i
            value = str(source_columns.get(key, ""))
            ws.cell(row=row, column=col_idx, value=value)

        # Round 1 sub‑columns
        # Tester (col round1_start+0) empty
        # Test Data (round1_start+1)
        ws.cell(row=row, column=round1_start + 1, value=test_data)
        # Test Results (round1_start+2)
        ws.cell(row=row, column=round1_start + 2, value="Todo")
        # Error Type (round1_start+3) empty
        # Error Description (round1_start+4) empty
        # Call ID (round1_start+5) empty
        # FPT Comment (round1_start+6) empty

        # Round 2 sub‑columns
        # Tester (col round2_start+0) empty
        # Test Data (round2_start+1)
        ws.cell(row=row, column=round2_start + 1, value=test_data_round2)
        # Test Results (round2_start+2)
        ws.cell(row=row, column=round2_start + 2, value="Todo")
        # Error Type (round2_start+3) empty
        # Error Description (round2_start+4) empty
        # Call ID (round2_start+5) empty
        # FPT Comment (round2_start+6) empty

        # Apply group fill **only to columns A-G** (fixed part + Expected Action Code)
        if fill is not None:
            for c in range(1, 8):
                ws.cell(row=row, column=c).fill = fill

        return row + 1

    # ---------- Write test cases ----------
    section_header_rows: list[int] = []

    if not group_by_step:
        row = 3   # data starts at row 3 because rows 1-2 are headers
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
        row = 3

        for step_no in ordered_groups:
            step_name = step_name_map.get(step_no, step_no)
            section_label = f"{step_no} - {step_name}"

            # Section header row – merge across all columns (1 to total_columns)
            ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=total_columns)
            cell_sec = ws.cell(row=row, column=1, value=section_label)
            cell_sec.font = SECTION_FONT
            cell_sec.fill = SECTION_FILL
            section_header_rows.append(row)
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
    _set_column_widths(ws, dynamic_keys)
    _apply_alignments(ws, dynamic_keys, section_header_rows=section_header_rows)
    _auto_fit_row_heights(ws)
    _add_data_validations(ws, dynamic_keys, ws.max_row)
    _add_conditional_formatting(ws, dynamic_keys, ws.max_row)
    _add_input_rows_sheet(wb, source_rows or [])

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
    source_rows = _load_source_rows(args.rows_path)
    export_to_excel(
        cases,
        args.out_path,
        step_name_map=step_name_map,
        group_by_step=args.group_by_step,
        use_group_fills=args.use_group_fills,
        allow_highlight_last=args.allow_highlight_last,
        source_rows=source_rows,
    )
    print(f"Wrote {len(cases)} test cases to Excel: {args.out_path}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())