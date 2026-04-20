from __future__ import annotations

import argparse
import json
from pathlib import Path
from typing import Any

from excel_to_json import _pick_excel_file, convert_excel_rows_to_json
from tc_to_excel import export_to_excel
import tcgen_e2e_human
import tcgen_multi_responses
import tcgen_output_human


GeneratorModule = Any


def _serialize_cases(cases: list[dict[str, Any]]) -> list[dict[str, Any]]:
    payload: list[dict[str, Any]] = []
    for i, tc in enumerate(cases):
        bot_responses_str = []
        for resp_item in tc.get("bot_responses", []):
            if isinstance(resp_item, str):
                bot_responses_str.append(resp_item)
            elif resp_item:
                bot_responses_str.append(" \\ ".join(resp_item))
            else:
                bot_responses_str.append("")
        payload.append({
            "tc_id": f"TC{i+1:03d}",
            "conditions": tc.get("conditions", ""),
            "steps": tc["steps"],
            "bot_responses": bot_responses_str,
            "expected_action_code": tc.get("expected_action_code", "N/A"),
            "path": tc.get("path", ""),
            "highlight_last_step": bool(tc.get("highlight_last_step", False)),
        })
    return payload


def _detect_root(rows: list[dict[str, str]]) -> str:
    sources = []
    source_set = set()
    dest_set = set()

    for row in rows:
        src = str(row.get("step_no", "")).strip()
        dst = str(row.get("next_step", "")).strip()
        if src and src not in source_set:
            sources.append(src)
            source_set.add(src)
        if dst and dst.lower() not in {"end", "stop", "__end__", "__terminal__"}:
            dest_set.add(dst)

    root_candidates = [src for src in sources if src not in dest_set]
    if root_candidates:
        return root_candidates[0]
    if sources:
        return sources[0]
    raise ValueError("Could not detect a root step from the Excel rows.")


def run_pipeline(
    *,
    mode: str,
    excel_path: Path | None,
    sheet: str | int | None,
    root: str | None,
    max_depth: int,
    rows_out: Path,
    testcases_out: Path,
    excel_out: Path,
) -> tuple[Path, Path, Path, int, str]:
    generator: GeneratorModule
    if mode == "e2e":
        generator = tcgen_e2e_human
    elif mode == "multi_responses":
        generator = tcgen_multi_responses
    elif mode == "output":
        generator = tcgen_output_human
    else:
        raise ValueError(f"Unsupported mode: {mode}")

    if excel_path is None:
        excel_path = _pick_excel_file(Path(__file__).resolve().parent / "input")

    rows = convert_excel_rows_to_json(excel_path, sheet=sheet)
    rows_out.parent.mkdir(parents=True, exist_ok=True)
    rows_out.write_text(json.dumps(rows, ensure_ascii=False, indent=2), encoding="utf-8")

    effective_root = root or _detect_root(rows)

    graph = generator.build_graph(rows)
    cases = generator.generate_test_cases(graph, root=effective_root, max_depth=max_depth)
    serialized_cases = _serialize_cases(cases)

    testcases_out.parent.mkdir(parents=True, exist_ok=True)
    testcases_out.write_text(
        json.dumps(serialized_cases, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )

    export_to_excel(serialized_cases, excel_out)
    return rows_out, testcases_out, excel_out, len(serialized_cases), effective_root


def main() -> int:
    parser = argparse.ArgumentParser(
        description="Run Excel -> JSON -> testcase generation -> Excel export in one command."
    )
    parser.add_argument("mode", choices=["e2e", "output", "multi_responses"])
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
        "--root",
        type=str,
        default=None,
        help="Root step. If omitted, auto-detects the first root-like step from the sheet.",
    )
    parser.add_argument("--max-depth", type=int, default=200)
    parser.add_argument(
        "--rows-out",
        type=Path,
        default=None,
        help="Output JSON path for normalized rows.",
    )
    parser.add_argument(
        "--testcases-out",
        type=Path,
        default=None,
        help="Output JSON path for generated test cases.",
    )
    parser.add_argument(
        "--excel-out",
        type=Path,
        default=None,
        help="Output Excel path for generated test cases.",
    )
    args = parser.parse_args()

    sheet = int(args.sheet) if isinstance(args.sheet, str) and args.sheet.isdigit() else args.sheet

    default_rows_out = Path(f"output/rows_{args.mode}.json")
    default_testcases_out = Path(f"output/testcases_{args.mode}.json")
    default_excel_out = Path(f"output/testcases_{args.mode}.xlsx")

    rows_out, testcases_out, excel_out, count, effective_root = run_pipeline(
        mode=args.mode,
        excel_path=args.file,
        sheet=sheet,
        root=args.root,
        max_depth=args.max_depth,
        rows_out=args.rows_out or default_rows_out,
        testcases_out=args.testcases_out or default_testcases_out,
        excel_out=args.excel_out or default_excel_out,
    )

    print(f"Pipeline mode: {args.mode}")
    print(f"Root: {effective_root}")
    print(f"Rows JSON: {rows_out}")
    print(f"Testcases JSON: {testcases_out}")
    print(f"Excel: {excel_out}")
    print(f"Generated {count} test cases")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
