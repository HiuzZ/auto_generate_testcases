from __future__ import annotations

import argparse
import ast
import json
import random
import re
from pathlib import Path
from typing import Any

from excel_to_json import _pick_excel_file, convert_excel_rows_to_json
from tc_to_excel import export_to_excel

# Import test case generators
import tcgen_e2e_human
import tcgen_e2e_short
import tcgen_multi_responses
import tcgen_output_human
import tcgen_output_short

try:
    import tcgen_e2e_short_v2
except ImportError:
    tcgen_e2e_short_v2 = None

# Import hybrid test data generator (if available)
try:
    from generate_test_data_hybrid import _generate_hybrid
except ImportError:
    _generate_hybrid = None

GeneratorModule = Any

_COND_SPLIT_RE = re.compile(r"\s*\\\s*|\n+")


def _normalize_conditions_text(text: str) -> str:
    parts = [part.strip() for part in _COND_SPLIT_RE.split(str(text)) if part.strip()]
    deduped: list[str] = []
    seen: set[str] = set()
    for part in parts:
        if part in seen:
            continue
        seen.add(part)
        deduped.append(part)
    timing_tokens = {"Business Hour", "Out of business hour"}
    timing_parts = [part for part in deduped if part in timing_tokens]
    base_parts = [part for part in deduped if part not in timing_tokens]
    base_text = " \\ ".join(base_parts)
    timing_text = " \\ ".join(timing_parts)
    if base_text and timing_text:
        return f"{base_text}\n\n{timing_text}"
    return base_text or timing_text


def _build_response_count_map(rows: list[dict[str, str]]) -> dict[tuple[str, str, str], int]:
    mapping: dict[tuple[str, str, str], int] = {}
    for row in rows:
        step_no = str(row.get("step_no", "")).strip()
        intent = str(row.get("customer_intent", "")).strip()
        condition = _normalize_conditions_text(str(row.get("conditions", "")))
        if not step_no or not intent:
            continue
        count = 0
        for key in ["bot_response", "bot_response_2", "bot_response_3", "bot_response_4", "bot_response_5"]:
            if str(row.get(key, "")).strip():
                count += 1
        if count <= 0:
            continue
        map_key = (step_no, intent, condition)
        mapping[map_key] = max(mapping.get(map_key, 0), count)
    return mapping


def _serialize_cases(
    cases: list[dict[str, Any]],
    *,
    response_count_map: dict[tuple[str, str, str], int] | None = None,
) -> list[dict[str, Any]]:
    payload: list[dict[str, Any]] = []
    seen_signatures: set[tuple[Any, ...]] = set()
    for i, tc in enumerate(cases):
        path_nodes = [part.strip() for part in str(tc.get("path", "")).split("->") if part.strip()]
        steps = list(tc.get("steps", []))
        case_conditions = {part.strip() for part in _COND_SPLIT_RE.split(str(tc.get("conditions", ""))) if part.strip()}

        def _match_random_count(step_idx: int, step_text: str) -> int:
            if not response_count_map or step_idx >= len(path_nodes) - 1:
                return 0
            src = path_nodes[step_idx]
            intents = [part.strip() for part in str(step_text).split(" \\ ") if part.strip()]
            best = 0
            for intent in intents:
                for (m_step, m_intent, m_condition), m_count in response_count_map.items():
                    if m_step != src or m_intent != intent:
                        continue
                    if m_condition and m_condition not in case_conditions:
                        continue
                    best = max(best, m_count)
            return best

        bot_responses_str = []
        for idx, resp_item in enumerate(tc.get("bot_responses", [])):
            if isinstance(resp_item, str):
                normalized_resp = resp_item.strip()
                if normalized_resp.startswith("[") and normalized_resp.endswith("]"):
                    try:
                        parsed = ast.literal_eval(normalized_resp)
                    except Exception:
                        parsed = None
                    if isinstance(parsed, list):
                        bot_responses_str.append(" \\ ".join(str(x) for x in parsed))
                    else:
                        bot_responses_str.append(resp_item)
                else:
                    bot_responses_str.append(resp_item)
            elif resp_item:
                responses = [str(x) for x in resp_item if str(x)]
                if not responses:
                    bot_responses_str.append("")
                elif len(responses) == 1:
                    bot_responses_str.append(responses[0])
                else:
                    bot_responses_str.append(f"{responses[0]} (random {len(responses)})")
            else:
                bot_responses_str.append("")

            if idx < len(bot_responses_str) and idx < len(steps):
                text = bot_responses_str[idx]
                if text and "(random " not in text:
                    random_count = _match_random_count(idx, str(steps[idx]))
                    if random_count > 1:
                        bot_responses_str[idx] = f"{text} (random {random_count})"
        normalized_conditions = _normalize_conditions_text(str(tc.get("conditions", "")))
        signature = (
            normalized_conditions,
            tuple(str(step) for step in tc.get("steps", [])),
            tuple(bot_responses_str),
            str(tc.get("expected_action_code", "N/A")),
            str(tc.get("path", "")),
            bool(tc.get("highlight_last_step", False)),
        )
        if signature in seen_signatures:
            continue
        seen_signatures.add(signature)
        payload.append({
            "tc_id": f"TC{len(payload)+1:03d}",
            "conditions": normalized_conditions,
            "steps": tc["steps"],
            "bot_responses": bot_responses_str,
            "expected_action_code": tc.get("expected_action_code", "N/A"),
            "path": tc.get("path", ""),
            "highlight_last_step": bool(tc.get("highlight_last_step", False)),
            "test_data": "",  # filled later if --gen-data is used
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

    if "A0" in source_set and "A1" in source_set:
        return "A1"

    root_candidates = [src for src in sources if src not in dest_set]
    if root_candidates:
        return root_candidates[0]
    if sources:
        return sources[0]
    raise ValueError("Could not detect a root step from the Excel rows.")


def _build_step_name_map(rows: list[dict[str, str]]) -> dict[str, str]:
    mapping: dict[str, str] = {}
    for row in rows:
        step_no = str(row.get("step_no", "")).strip()
        step_name = str(row.get("step_name", "")).strip()
        if not step_no or not step_name:
            continue
        if step_name.upper() == step_no.upper():
            continue
        if step_no not in mapping:
            mapping[step_no] = step_name
    return mapping


def _generate_test_data_for_cases(serialized_cases: list[dict[str, Any]]) -> None:
    """Fill in the 'test_data' field using the hybrid generator (TLS + Llama)."""
    if _generate_hybrid is None:
        print("⚠️  Warning: generate_test_data_hybrid not available. Test Data will remain empty.")
        return

    for case in serialized_cases:
        steps = case["steps"]
        if not steps:
            # A0 or empty case – leave test_data blank
            continue

        # Convert "resp1 \\ resp2" strings into a randomly chosen single response per step
        bot_responses_list: list[str] = []
        for resp_str in case["bot_responses"]:
            parts = [p.strip() for p in resp_str.split(" \\ ") if p.strip()]
            bot_responses_list.append(random.choice(parts) if parts else "")
        while len(bot_responses_list) < len(steps):
            bot_responses_list.append("")

        tc_id = case["tc_id"]
        test_data = _generate_hybrid(steps, bot_responses_list, tc_id)
        case["test_data"] = test_data
        print(f"  Generated test data for {tc_id}")


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
    gen_data: bool = False,
) -> tuple[Path, Path, Path, int, str]:
    generator: GeneratorModule
    if mode == "e2e":
        generator = tcgen_e2e_human
    elif mode == "e2e_short":
        generator = tcgen_e2e_short
    elif mode == "e2e_short_v2":
        if tcgen_e2e_short_v2 is None:
            raise ValueError("Unsupported mode: e2e_short_v2 (tcgen_e2e_short_v2.py not found)")
        generator = tcgen_e2e_short_v2
    elif mode == "multi_responses":
        generator = tcgen_multi_responses
    elif mode == "output":
        generator = tcgen_output_human
    elif mode == "output_short":
        generator = tcgen_output_short
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
    response_count_map = _build_response_count_map(rows) if mode in {"e2e", "e2e_short", "e2e_short_v2", "output", "output_short"} else None
    serialized_cases = _serialize_cases(cases, response_count_map=response_count_map)

    if gen_data:
        print("\n--- Generating test data (hybrid) ---")
        _generate_test_data_for_cases(serialized_cases)

    testcases_out.parent.mkdir(parents=True, exist_ok=True)
    testcases_out.write_text(
        json.dumps(serialized_cases, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )

    step_name_map = _build_step_name_map(rows)
    # Grouping by step is enabled for all pipelines,
    # but group fills and red highlighting are only for multi_responses.
    group_by_step = True
    use_group_fills = (mode == "multi_responses")
    allow_highlight_last = (mode == "multi_responses")

    export_to_excel(
        serialized_cases,
        excel_out,
        step_name_map=step_name_map,
        group_by_step=group_by_step,
        use_group_fills=use_group_fills,
        allow_highlight_last=allow_highlight_last,
    )
    return rows_out, testcases_out, excel_out, len(serialized_cases), effective_root


def main() -> int:
    parser = argparse.ArgumentParser(
        description="Run Excel -> JSON -> testcase generation -> Excel export in one command."
    )
    parser.add_argument("mode", choices=["e2e", "e2e_short", "e2e_short_v2", "output", "output_short", "multi_responses"])
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
    parser.add_argument(
        "--gen-data",
        action="store_true",
        default=False,
        help="Generate customer utterances (hybrid) and fill Test Data column.",
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
        gen_data=args.gen_data,
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
