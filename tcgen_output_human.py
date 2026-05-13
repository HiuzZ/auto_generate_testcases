from __future__ import annotations

import argparse
import csv
import json
import re
from collections import defaultdict
from dataclasses import dataclass
from pathlib import Path
from typing import Any, DefaultDict, Iterable, Dict, List, Tuple, Set, Optional

import tcgen_e2e_human


@dataclass(frozen=True)
class Transition:
    src: str
    dst: str
    condition: str
    intent: str
    action_code: str
    bot_response: str
    source_row: dict[str, str]


_REPEAT_RE = re.compile(r"(?i)\blần\s*(\d+)\b")


def _clean_str(v: Any) -> str:
    if v is None:
        return ""
    s = str(v).strip()
    return "" if s.lower() == "nan" else s


PERMANENT_ROW_KEYS = tcgen_e2e_human.PERMANENT_ROW_KEYS
_clean_row = tcgen_e2e_human._clean_row
_source_columns = tcgen_e2e_human._source_columns
_render_bot_response_item = tcgen_e2e_human._render_bot_response_item


def _is_terminal_step(step_no: str) -> bool:
    s = step_no.strip().lower()
    return s in {"end", "stop", "__end__", "__terminal__"}


def _parse_repeat(intent: str) -> tuple[str, int] | None:
    m = _REPEAT_RE.search(intent)
    if not m:
        return None
    n = int(m.group(1))
    base = _REPEAT_RE.sub("", intent)
    base = " ".join(base.strip().split())
    return base, n


def load_transitions_json(path: Path) -> list[dict[str, str]]:
    data = json.loads(path.read_text(encoding="utf-8"))
    if not isinstance(data, list):
        raise ValueError("Input must be a JSON list of transition objects.")
    rows: list[dict[str, str]] = []
    for i, obj in enumerate(data):
        if not isinstance(obj, dict):
            raise ValueError(f"Item at index {i} is not an object.")
        row = _clean_row(obj)
        for key in PERMANENT_ROW_KEYS:
            row.setdefault(key, "")
        rows.append(row)
    return rows


def build_graph(
    transitions: Iterable[dict[str, str]],
) -> dict[str, list[Transition]]:
    adjacency: DefaultDict[str, list[Transition]] = defaultdict(list)
    for t in transitions:
        src = _clean_str(t.get("step_no", ""))
        dst = _clean_str(t.get("next_step", ""))
        if not dst:
            dst = "End"
        condition = _clean_str(t.get("conditions", ""))
        intent = _clean_str(t.get("customer_intent", ""))
        if not intent:
            intent = "(empty intent)"
        action_code = _clean_str(t.get("action_code", ""))
        if not action_code:
            action_code = "N/A"
        bot_response = _clean_str(t.get("bot_response", ""))
        if not src:
            continue
        source_row = dict(t)
        source_row.update({
            "step_no": src,
            "step_name": _clean_str(t.get("step_name", "")),
            "conditions": condition,
            "customer_intent": intent,
            "bot_response": bot_response,
            "next_step": dst,
            "action_code": action_code,
        })
        adjacency[src].append(
            Transition(
                src=src,
                dst=dst,
                condition=condition,
                intent=intent,
                action_code=action_code,
                bot_response=bot_response,
                source_row=source_row,
            )
        )
    for src in adjacency:
        adjacency[src].sort(key=lambda tr: (tr.dst, tr.action_code, tr.condition, tr.intent))
    return dict(adjacency)


def generate_test_cases(
    graph: dict[str, list[Transition]],
    *,
    root: str = "A1",
    max_depth: int = 200,
) -> list[dict[str, Any]]:
    """
    Generate test cases by DFS over the transition graph.

    Key rules:
    - Non-repeat intents: keep each source row as its own transition.
    - Repeat intents (lần N): each base+chain is kept SEPARATE from non-repeat
      and from other bases. Walk lần 1 → lần 2 → lần 3 in sequence.
    - visited_edges prevents re-traversing the same (src, dst, intent) edge.
    """
    cases: List[Dict[str, Any]] = []

    def _emit(steps, bot_responses_list, action_code, path_nodes, source_rows):
        cases.append({
            "steps": list(steps),
            "bot_responses": [list(r) for r in bot_responses_list],
            "expected_action_code": action_code,
            "path": " -> ".join(path_nodes),
            "source_rows": list(source_rows),
        })

    def _make_step(trans_list: List[Transition]) -> tuple[str, List[str], str]:
        intents = sorted(set(tr.intent for tr in trans_list))
        step_intent = " \\ ".join(intents)
        bot_resp = list(dict.fromkeys(tr.bot_response for tr in trans_list if tr.bot_response))
        conditions = list(dict.fromkeys(tr.condition for tr in trans_list))
        step_condition = conditions[0] if conditions else ""
        return step_intent, bot_resp, step_condition

    def _make_single_step(tr: Transition) -> tuple[str, List[str], str]:
        bot_resp = [tr.bot_response] if tr.bot_response else []
        return tr.intent, bot_resp, tr.condition

    def _append_condition(case_conditions: List[str], condition: str) -> List[str]:
        if condition in case_conditions:
            return list(case_conditions)
        return case_conditions + [condition]

    def _render_conditions(case_conditions: List[str]) -> str:
        non_empty = [cond for cond in case_conditions if cond]
        if non_empty:
            return " \\ ".join(non_empty)
        return ""

    def dfs(
        node: str,
        steps: List[str],
        bot_responses_list: List[List[str]],
        case_conditions: List[str],
        path_nodes: List[str],
        source_rows: List[dict[str, str]],
        visited_edges: Set[Tuple[str, str, str, str]],  # (src, dst, intent, condition)
        consumed_chains: Set[Tuple[str, str]],
    ) -> None:
        if len(steps) >= max_depth:
            return

        transitions = graph.get(node, [])
        if not transitions:
            return

        repeat_map: Dict[str, Dict[int, Dict[Tuple[str, str, str], List[Transition]]]] = defaultdict(
            lambda: defaultdict(lambda: defaultdict(list))
        )
        normal_transitions: List[Transition] = []

        for tr in transitions:
            rep = _parse_repeat(tr.intent)
            if rep:
                base, n = rep
                repeat_map[base][n][(tr.dst, tr.action_code, tr.condition)].append(tr)
            else:
                normal_transitions.append(tr)

        # ── 1. Process normal transitions (no intent grouping) ────────────
        for tr in sorted(normal_transitions, key=lambda tr: (tr.dst, tr.action_code, tr.condition, tr.intent)):
            dst = tr.dst
            action_code = tr.action_code
            step_intent, bot_resp, step_condition = _make_single_step(tr)
            edge_key = (node, dst, step_intent, step_condition)

            if edge_key in visited_edges:
                continue

            new_steps = steps + [step_intent]
            new_bot = bot_responses_list + [bot_resp]
            new_conditions = _append_condition(case_conditions, step_condition)
            new_path = path_nodes + [dst]
            new_source_rows = source_rows + [tr.source_row]
            new_visited = visited_edges | {edge_key}

            _emit(new_steps, new_bot, action_code, new_path, new_source_rows)
            cases[-1]["conditions"] = _render_conditions(new_conditions)

            if not _is_terminal_step(dst) and dst in graph:
                dfs(dst, new_steps, new_bot, new_conditions, new_path, new_source_rows, new_visited, consumed_chains)

        # ── 2. Process ordered repeat chains per base ──────────────────────
        for base in sorted(repeat_map.keys()):
            n_map = repeat_map[base]
            if 1 not in n_map:
                continue

            chain_id = {("__repeat_base__", base)}
            if chain_id & consumed_chains:
                continue

            chain_levels: List[Tuple[int, List[Tuple[str, str, str, List[Transition]]]]] = []
            expected_n = 1
            while expected_n in n_map:
                grouped = [
                    (dst, action_code, step_condition, trans_list)
                    for (dst, action_code, step_condition), trans_list in sorted(n_map[expected_n].items())
                ]
                chain_levels.append((expected_n, grouped))
                expected_n += 1

            _walk_repeat_chain(
                origin_node=node,
                chain_ids=chain_id,
                chain_levels=chain_levels,
                chain_idx=0,
                steps=steps,
                bot_responses_list=bot_responses_list,
                case_conditions=case_conditions,
                path_nodes=path_nodes,
                source_rows=source_rows,
                visited_edges=visited_edges,
                consumed_chains=consumed_chains,
            )

    def _walk_repeat_chain(
        origin_node: str,
        chain_ids: Set[Tuple[str, str]],
        chain_levels: List[Tuple[int, List[Tuple[str, str, str, List[Transition]]]]],
        chain_idx: int,
        steps: List[str],
        bot_responses_list: List[List[str]],
        case_conditions: List[str],
        path_nodes: List[str],
        source_rows: List[dict[str, str]],
        visited_edges: Set[Tuple[str, str, str, str]],
        consumed_chains: Set[Tuple[str, str]],
    ) -> None:
        if chain_idx >= len(chain_levels):
            return
        if len(steps) >= max_depth:
            return

        _, grouped_transitions = chain_levels[chain_idx]
        is_last = chain_idx == len(chain_levels) - 1

        for dst, action_code, step_condition, trans_list in grouped_transitions:
            step_intent, bot_resp, step_condition = _make_step(trans_list)
            edge_key = (origin_node, dst, step_intent, step_condition)

            if edge_key in visited_edges:
                continue

            new_steps = steps + [step_intent]
            new_bot = bot_responses_list + [bot_resp]
            new_conditions = _append_condition(case_conditions, step_condition)
            new_path = path_nodes + [dst]
            new_visited = visited_edges | {edge_key}
            new_source_rows = source_rows + [tr.source_row for tr in trans_list]
            new_consumed = consumed_chains | chain_ids

            if _is_terminal_step(dst):
                _emit(new_steps, new_bot, action_code, new_path, new_source_rows)
                cases[-1]["conditions"] = _render_conditions(new_conditions)
                continue

            if is_last:
                _emit(new_steps, new_bot, action_code, new_path, new_source_rows)
                cases[-1]["conditions"] = _render_conditions(new_conditions)
                if dst in graph:
                    dfs(dst, new_steps, new_bot, new_conditions, new_path, new_source_rows, new_visited, new_consumed)
            else:
                _walk_repeat_chain(
                    origin_node=origin_node,
                    chain_ids=chain_ids,
                    chain_levels=chain_levels,
                    chain_idx=chain_idx + 1,
                    steps=new_steps,
                    bot_responses_list=new_bot,
                    case_conditions=new_conditions,
                    path_nodes=new_path,
                    source_rows=new_source_rows,
                    visited_edges=new_visited,
                    consumed_chains=new_consumed,
                )

    dfs(root, [], [], [], [root], [], set(), set())

    # Deduplicate
    seen: Set[Tuple] = set()
    unique: List[Dict[str, Any]] = []
    for case in cases:
        key = (tuple(case["steps"]), case.get("conditions", ""), case["expected_action_code"])
        if key not in seen:
            seen.add(key)
            unique.append(case)

    if root != "A0":
        a0_case = tcgen_e2e_human._a0_greeting_case(graph)
        if a0_case is not None:
            unique = [a0_case] + unique
    return unique


def write_cases_json(cases: list[dict[str, Any]], out_path: Path) -> None:
    out_path.parent.mkdir(parents=True, exist_ok=True)
    payload = []
    for i, tc in enumerate(cases):
        bot_responses_str = [_render_bot_response_item(resp) for resp in tc.get("bot_responses", [])]
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
    out_path.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")


def write_cases_csv(cases: list[dict[str, Any]], out_path: Path) -> None:
    out_path.parent.mkdir(parents=True, exist_ok=True)
    max_len = max((len(tc["steps"]) for tc in cases), default=0)
    headers = [
        "tc_id", "conditions", "path", "expected_action_code",
        *[f"step_{i+1}" for i in range(max_len)],
        *[f"bot_response_{i+1}" for i in range(max_len)],
    ]
    with out_path.open("w", encoding="utf-8", newline="") as f:
        w = csv.writer(f)
        w.writerow(headers)
        for i, tc in enumerate(cases):
            steps = tc["steps"]
            bot_responses_raw = tc.get("bot_responses", [])
            bot_responses = [
                " \\ ".join(r) if r else ""
                for r in bot_responses_raw
            ]
            conditions = tc.get("conditions", "")
            action_code = tc.get("expected_action_code", "N/A")
            path = str(tc.get("path", ""))
            padded_steps = steps + [""] * (max_len - len(steps))
            padded_resp = bot_responses + [""] * (max_len - len(bot_responses))
            w.writerow([f"TC{i+1:03d}", conditions, path, action_code] + padded_steps + padded_resp)


def main() -> int:
    parser = argparse.ArgumentParser(
        description="Generate Test Cases from decision tree dataset."
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
            bot_responses_str = [_render_bot_response_item(resp) for resp in tc.get("bot_responses", [])]
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
