from __future__ import annotations

import argparse
import json
from pathlib import Path
from typing import Any

import tcgen_e2e_human


Transition = tcgen_e2e_human.Transition
PERMANENT_ROW_KEYS = tcgen_e2e_human.PERMANENT_ROW_KEYS

load_transitions_json = tcgen_e2e_human.load_transitions_json
build_graph = tcgen_e2e_human.build_graph
write_cases_json = tcgen_e2e_human.write_cases_json
write_cases_csv = tcgen_e2e_human.write_cases_csv
_source_columns = tcgen_e2e_human._source_columns
_render_bot_response_item = tcgen_e2e_human._render_bot_response_item


def generate_test_cases(
    graph: dict[str, list[Transition]],
    *,
    root: str = "A1",
    max_depth: int = 200,
) -> list[dict[str, Any]]:
    return tcgen_e2e_human.generate_test_cases(
        graph,
        root=root,
        max_depth=max_depth,
        emit_at_every_step=True,
    )


def main() -> int:
    parser = argparse.ArgumentParser(
        description="Generate output test cases from decision tree dataset."
    )
    parser.add_argument("--in", dest="in_path", type=Path, required=True)
    parser.add_argument("--root", type=str, default="A1")
    parser.add_argument("--out", dest="out_path", type=Path, default=None)
    parser.add_argument("--format", choices=["json", "csv"], default="json")
    parser.add_argument("--max-depth", type=int, default=200)
    args = parser.parse_args()

    transitions = load_transitions_json(args.in_path)
    graph = build_graph(transitions)
    cases = generate_test_cases(graph, root=args.root, max_depth=args.max_depth)

    if args.out_path is None:
        payload = []
        for i, tc in enumerate(cases):
            bot_responses_str = [
                _render_bot_response_item(resp)
                for resp in tc.get("bot_responses", [])
            ]
            source_rows = list(tc.get("source_rows", []))
            payload.append({
                "tc_id": f"TC{i+1:03d}",
                "conditions": tc.get("conditions", ""),
                "steps": tc["steps"],
                "bot_responses": bot_responses_str,
                "expected_action_code": tc.get("expected_action_code", "N/A"),
                "path": tc.get("path", ""),
                "source_columns": _source_columns(source_rows),
            })
        print(json.dumps(payload, ensure_ascii=False, indent=2))
        return 0

    if args.format == "json":
        write_cases_json(cases, args.out_path)
    else:
        write_cases_csv(cases, args.out_path)

    print(f"Wrote {len(cases)} test cases to: {args.out_path}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
