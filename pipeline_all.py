from __future__ import annotations

import argparse
from pathlib import Path

import openpyxl

from excel_to_json import _pick_excel_file, read_data_schema_sheet
from pipeline_tc import run_pipeline
from tc_to_excel import _add_checklist_sheet, _add_data_input_sheet, _add_data_schema_sheet

try:
    from generate_test_data_hybrid import _generate_hybrid as _hybrid_available
except ImportError:
    _hybrid_available = None


# (mode, excel sheet name)
MODES_CONFIG = [
    ("e2e_short",        "Test case"),
    ("output_short",     "Test Output"),
    ("multi_responses",  "Test case Đa Thoại"),
]


def main() -> int:
    parser = argparse.ArgumentParser(
        description=(
            "Run e2e_short + output_short + multi_responses pipelines and export to a "
            "single Excel file. Test data is generated via generate_test_data_hybrid "
            "(TLS bot + Llama fallback) by default."
        )
    )
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
        help="Root step. If omitted, auto-detects from the sheet.",
    )
    parser.add_argument("--max-depth", type=int, default=200)
    parser.add_argument(
        "--suffix",
        type=str,
        default=None,
        help="Optional suffix for output filenames, e.g. 'vb' -> testcases_all_vb.xlsx",
    )
    parser.add_argument(
        "--no-gen-data",
        action="store_true",
        default=False,
        help="Skip Test Data generation (faster, useful for dry runs).",
    )
    args = parser.parse_args()

    gen_data = not args.no_gen_data
    if gen_data and _hybrid_available is None:
        print("⚠️  generate_test_data_hybrid not available — Test Data column will be empty.")
        gen_data = False

    sheet = int(args.sheet) if isinstance(args.sheet, str) and args.sheet.isdigit() else args.sheet
    suffix = f"_{args.suffix}" if args.suffix else ""

    # Resolve excel path early so we can read the schema sheet before the loop
    resolved_excel_path = (
        args.file
        if args.file is not None
        else _pick_excel_file(Path(__file__).resolve().parent / "input")
    )
    schema_keys, schema_value_rows, schema_all_rows = read_data_schema_sheet(resolved_excel_path)
    if schema_keys:
        print(f"Data schema: {len(schema_keys)} key(s): {schema_keys}")
    else:
        print("Data schema: not found in input Excel (using defaults)")

    print(f"Test data generation (hybrid): {'ON' if gen_data else 'OFF'}")

    combined_excel_out = Path(f"output/testcases_all{suffix}.xlsx")

    # Single workbook shared across all pipeline modes
    combined_wb = openpyxl.Workbook()
    combined_wb.remove(combined_wb.active)  # remove default empty sheet
    _add_checklist_sheet(combined_wb)       # "Checklist" is always the first sheet

    results: list[tuple[str, int, str]] = []

    for mode, tc_sheet_name in MODES_CONFIG:
        rows_out     = Path(f"output/rows_{mode}{suffix}.json")
        testcases_out = Path(f"output/testcases_{mode}{suffix}.json")
        # excel_out is required by run_pipeline signature but won't be saved
        # when combined_wb is provided
        excel_out = Path(f"output/testcases_{mode}{suffix}.xlsx")

        print(f"\n{'='*60}")
        print(f"Running pipeline: {mode}  →  sheet '{tc_sheet_name}'")
        print(f"{'='*60}")

        try:
            _, _, _, count, effective_root = run_pipeline(
                mode=mode,
                excel_path=resolved_excel_path,
                sheet=sheet,
                root=args.root,
                max_depth=args.max_depth,
                rows_out=rows_out,
                testcases_out=testcases_out,
                excel_out=excel_out,
                gen_data=gen_data,
                combined_wb=combined_wb,
                tc_sheet_name=tc_sheet_name,
            )
            results.append((mode, count, tc_sheet_name))
            print(f"Root      : {effective_root}")
            print(f"Rows JSON : {rows_out}")
            print(f"Cases JSON: {testcases_out}")
            print(f"Generated {count} test cases  →  sheet '{tc_sheet_name}'")
        except Exception as exc:
            print(f"ERROR in mode '{mode}': {exc}")
            results.append((mode, -1, tc_sheet_name))

    # Add shared helper sheets once, then save
    _add_data_input_sheet(
        combined_wb,
        schema_keys=schema_keys if schema_keys else None,
        schema_value_rows=schema_value_rows if schema_value_rows else None,
    )
    _add_data_schema_sheet(
        combined_wb,
        schema_all_rows=schema_all_rows if schema_all_rows else None,
    )
    combined_excel_out.parent.mkdir(parents=True, exist_ok=True)
    combined_wb.save(combined_excel_out)

    print(f"\n{'='*60}")
    print("Summary")
    print(f"{'='*60}")
    for mode, count, sheet_name in results:
        status = f"{count:4d} test cases" if count >= 0 else "     FAILED"
        print(f"  {mode:20s}  {status}  →  sheet '{sheet_name}'")
    print(f"\nCombined Excel: {combined_excel_out}")

    return 0 if all(count >= 0 for _, count, _ in results) else 1


if __name__ == "__main__":
    raise SystemExit(main())
