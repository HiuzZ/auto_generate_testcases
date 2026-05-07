from __future__ import annotations

import argparse
import json
from pathlib import Path
from typing import Any

import tcgen_e2e_human
import tcgen_e2e_short


Transition = tcgen_e2e_human.Transition
TransitionKey = tuple[str, str, str, str, str, str]


def _is_terminal_step(step_no: str) -> bool:
    return tcgen_e2e_short._is_terminal_step(step_no)


def _transition_key(tr: Transition) -> TransitionKey:
    dst = str(tr.dst).strip()
    action = str(tr.action_code).strip() if _is_terminal_step(dst) else ""
    return (
        str(tr.src).strip(),
        str(tr.condition).strip(),
        str(tr.intent).strip(),
        dst,
        action,
        str(tr.bot_response).strip(),
    )


def _dedupe_graph_for_representative_traversal(
    graph: dict[str, list[Transition]],
) -> tuple[dict[str, list[Transition]], set[TransitionKey]]:
    deduped: dict[str, list[Transition]] = {}
    skipped_keys: set[TransitionKey] = set()

    for src, transitions in graph.items():
        selected: list[Transition] = []
        seen_forward_groups: set[tuple[str, str]] = set()
        for tr in transitions:
            dst = str(tr.dst).strip()
            if str(src).strip() != dst:
                group_key = (str(src).strip(), dst)
                if group_key in seen_forward_groups:
                    skipped_keys.add(_transition_key(tr))
                    continue
                seen_forward_groups.add(group_key)
            selected.append(tr)
        deduped[src] = selected

    return deduped, skipped_keys


def _case_transition_keys(
    case: dict[str, Any],
    graph: dict[str, list[Transition]],
) -> set[TransitionKey]:
    path_nodes = tcgen_e2e_short._extract_nodes_from_path(str(case.get("path", "")))
    steps = [str(step) for step in case.get("steps", [])]
    bot_responses = list(case.get("bot_responses", []))
    case_conditions = tcgen_e2e_short._split_conditions(str(case.get("conditions", "")))
    expected_action = str(case.get("expected_action_code", "")).strip()
    covered: set[TransitionKey] = set()

    for idx, step_text in enumerate(steps):
        if idx >= len(path_nodes) - 1:
            break
        src = path_nodes[idx].strip()
        dst = path_nodes[idx + 1].strip()
        response_item = bot_responses[idx] if idx < len(bot_responses) else []
        if isinstance(response_item, str):
            response_values = {response_item.strip()} if response_item.strip() else set()
        else:
            response_values = {str(value).strip() for value in response_item if str(value).strip()}

        for intent in tcgen_e2e_short._split_step_intents(step_text):
            for tr in graph.get(src, []):
                if str(tr.dst).strip() != dst:
                    continue
                if str(tr.intent).strip() != intent:
                    continue
                tr_condition = str(tr.condition).strip()
                if tr_condition and tr_condition not in case_conditions:
                    continue
                if _is_terminal_step(dst) and expected_action and str(tr.action_code).strip() != expected_action:
                    continue
                tr_response = str(tr.bot_response).strip()
                if response_values and tr_response and tr_response not in response_values:
                    continue
                covered.add(_transition_key(tr))

    return covered


def _collect_transition_universe(graph: dict[str, list[Transition]]) -> set[TransitionKey]:
    return {_transition_key(tr) for transitions in graph.values() for tr in transitions}


def _expand_cases_by_full_transition_coverage(
    cases: list[dict[str, Any]],
    graph: dict[str, list[Transition]],
    traversal_graph: dict[str, list[Transition]],
    target_universe: set[TransitionKey],
    *,
    root: str,
    max_depth: int,
) -> list[dict[str, Any]]:
    universe = set(target_universe)
    if not universe:
        return cases

    selected = list(cases)
    selected_signatures = {tcgen_e2e_short._case_signature(case) for case in selected}
    covered: set[TransitionKey] = set()
    for case in selected:
        covered |= _case_transition_keys(case, graph)

    uncovered = universe - covered
    if not uncovered:
        return selected

    transition_by_key: dict[TransitionKey, Transition] = {}
    for transitions in graph.values():
        for tr in transitions:
            transition_by_key.setdefault(_transition_key(tr), tr)

    for missing_key in sorted(uncovered):
        if missing_key in covered:
            continue
        target = transition_by_key.get(missing_key)
        if target is None:
            continue

        resample_graph = {src: list(transitions) for src, transitions in traversal_graph.items()}
        target_src = str(target.src).strip()
        target_key = _transition_key(target)
        target_transitions = [
            tr for tr in resample_graph.get(target_src, []) if _transition_key(tr) != target_key
        ]
        resample_graph[target_src] = [target] + target_transitions

        candidates = []
        for idx, candidate in enumerate(tcgen_e2e_human.generate_test_cases(resample_graph, root=root, max_depth=max_depth)):
            candidate_keys = _case_transition_keys(candidate, graph)
            if missing_key not in candidate_keys:
                continue
            signature = tcgen_e2e_short._case_signature(candidate)
            if signature in selected_signatures:
                covered |= candidate_keys
                break
            candidates.append((idx, candidate, candidate_keys))

        if not candidates:
            continue

        chosen = min(
            candidates,
            key=lambda item: (
                tcgen_e2e_short._path_length(item[1]),
                len(item[1].get("steps", [])),
                item[0],
            ),
        )
        selected.append(chosen[1])
        selected_signatures.add(tcgen_e2e_short._case_signature(chosen[1]))
        covered |= chosen[2]

    return selected


def build_graph(transitions: list[dict[str, str]]) -> dict[str, list[Transition]]:
    return tcgen_e2e_human.build_graph(transitions)


def generate_test_cases(
    graph: dict[str, list[Transition]],
    *,
    root: str = "A1",
    max_depth: int = 200,
) -> list[dict[str, Any]]:
    traversal_graph, skipped_keys = _dedupe_graph_for_representative_traversal(graph)
    cases = tcgen_e2e_short.generate_test_cases(
        traversal_graph,
        root=root,
        max_depth=max_depth,
    )
    cases = _expand_cases_by_full_transition_coverage(
        cases,
        graph,
        traversal_graph,
        skipped_keys,
        root=root,
        max_depth=max_depth,
    )
    cases = tcgen_e2e_short._preserve_combined_conditions(cases)
    cases = tcgen_e2e_short._apply_tc051_tc052_hotfix(cases, graph)
    cases = tcgen_e2e_short._ensure_terminal_sibling_condition_variants(cases, graph)
    if root != "A0":
        return tcgen_e2e_human._prepend_single_a0_case(cases, graph, multi_response=False)
    return cases


def main() -> int:
    parser = argparse.ArgumentParser(
        description="Generate short E2E test cases with representative traversal and full row coverage."
    )
    parser.add_argument("--in", dest="in_path", type=Path, required=True)
    parser.add_argument("--root", type=str, default="A1")
    parser.add_argument("--out", dest="out_path", type=Path, default=None)
    parser.add_argument("--format", choices=["json", "csv"], default="json")
    parser.add_argument("--max-depth", type=int, default=200)
    args = parser.parse_args()

    transitions = tcgen_e2e_human.load_transitions_json(args.in_path)
    graph = build_graph(transitions)
    cases = generate_test_cases(graph, root=args.root, max_depth=args.max_depth)

    if args.out_path is None:
        payload = []
        for i, tc in enumerate(cases):
            bot_responses_str = [
                " \\ ".join(resp_list) if resp_list else ""
                for resp_list in tc.get("bot_responses", [])
            ]
            payload.append({
                "tc_id": f"TC{i+1:03d}",
                "conditions": tc.get("conditions", ""),
                "steps": tc["steps"],
                "bot_responses": bot_responses_str,
                "expected_action_code": tc.get("expected_action_code", "N/A"),
                "path": tc.get("path", ""),
            })
        print(json.dumps(payload, ensure_ascii=False, indent=2))
        return 0

    if args.format == "json":
        tcgen_e2e_human.write_cases_json(cases, args.out_path)
    else:
        tcgen_e2e_human.write_cases_csv(cases, args.out_path)

    print(f"Wrote {len(cases)} test cases to: {args.out_path}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
