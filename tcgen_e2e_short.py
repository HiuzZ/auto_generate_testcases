from __future__ import annotations

import argparse
import csv
import json
import re
from collections import defaultdict
from dataclasses import dataclass
from pathlib import Path
from typing import Any, DefaultDict, Iterable, Dict, List, Tuple, Set, Optional


PERMANENT_ROW_KEYS = [
    "step_no",
    "step_name",
    "conditions",
    "customer_intent",
    "bot_response",
    "bot_response_2",
    "bot_response_3",
    "bot_response_4",
    "bot_response_5",
    "next_step",
    "action_code",
]


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

# FIX 1: Use a non-greedy atomic group for the node-list part so the whole
# "YES_A7_A8_A9_A10_A11" token is captured as one unit.
# group(1) = YES or NO
# group(2) = A7_A8_A9_A10_A11  (then we split on "_" inside the function)
_SPECIAL_STEP_COND_RE = re.compile(
    r"\b(YES|NO)_([A-Za-z0-9]+(?:_[A-Za-z0-9]+)*)\b"
)


def _clean_str(v: Any) -> str:
    if v is None:
        return ""
    s = str(v).strip()
    return "" if s.lower() == "nan" else s


def _clean_row(obj: dict[str, Any]) -> dict[str, str]:
    return {str(k): _clean_str(v) for k, v in obj.items()}


def _source_columns(source_rows: list[dict[str, str]]) -> dict[str, str]:
    keys = list(PERMANENT_ROW_KEYS)
    for row in source_rows:
        for key in row:
            if key not in keys:
                keys.append(key)

    summary: dict[str, str] = {}
    for key in keys:
        values: list[str] = []
        seen_values: set[str] = set()
        for row in source_rows:
            value = _clean_str(row.get(key, ""))
            if not value or value in seen_values:
                continue
            values.append(value)
            seen_values.add(value)
        summary[key] = "\n".join(values)
    return summary


def _render_bot_response_item(resp_item: Any) -> str:
    if isinstance(resp_item, str):
        return resp_item
    if resp_item:
        return " \\ ".join(str(resp) for resp in resp_item if str(resp))
    return ""


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


def _evaluate_step_condition(
    raw_condition: str,
    traversed_nodes: set[str],
) -> tuple[bool, str]:
    """
    Evaluate YES_*/NO_* special tokens against the set of already-traversed nodes.

    YES_A7_A8_A9_A10_A11  →  allow only if the path passed through A7 OR A8 OR … OR A11
    NO_A7_A8_A9_A10_A11   →  allow only if the path did NOT pass through any of them

    Returns (allowed, cleaned_condition_string).
    The cleaned string has all YES_*/NO_* tokens removed so only the human-readable
    condition text remains for the output conditions column.
    """
    text = str(raw_condition or "")

    for match in _SPECIAL_STEP_COND_RE.finditer(text):
        mode = match.group(1).upper()
        # Split on underscore to recover individual node IDs.
        # filter(None, ...) drops empty strings from accidental double-underscores.
        nodes = [token.strip() for token in match.group(2).split("_") if token.strip()]
        if not nodes:
            continue

        passed_any = any(node in traversed_nodes for node in nodes)

        if mode == "YES" and not passed_any:
            return False, ""
        if mode == "NO" and passed_any:
            return False, ""

    # FIX 2: Strip the special tokens cleanly, then collapse any orphaned
    # separator fragments (e.g. a lone " \\ " with nothing on one side).
    cleaned = _SPECIAL_STEP_COND_RE.sub("", text)
    parts = [p.strip() for p in re.split(r"\s*\\\s*", cleaned) if p.strip()]
    return True, " \\ ".join(parts)


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
            "bot_response_2": _clean_str(t.get("bot_response_2", "")),
            "bot_response_3": _clean_str(t.get("bot_response_3", "")),
            "bot_response_4": _clean_str(t.get("bot_response_4", "")),
            "bot_response_5": _clean_str(t.get("bot_response_5", "")),
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


def _a0_greeting_case(
    graph: dict[str, list[Transition]],
    *,
    multi_response: bool = False,
) -> dict[str, Any] | None:
    transitions = graph.get("A0", [])
    if not transitions:
        return None

    first = transitions[0]
    if multi_response:
        bot_responses = [tr.bot_response for tr in transitions if tr.bot_response]
    else:
        bot_responses = [first.bot_response] if first.bot_response else []

    return {
        "conditions": "",
        "steps": [],
        "bot_responses": bot_responses,
        "expected_action_code": first.action_code,
        "path": "A0",
        "source_rows": [first.source_row],
    }


def _prepend_single_a0_case(
    cases: list[dict[str, Any]],
    graph: dict[str, list[Transition]],
    *,
    multi_response: bool = False,
) -> list[dict[str, Any]]:
    a0_case = _a0_greeting_case(graph, multi_response=multi_response)
    if a0_case is None:
        return cases
    non_a0_cases = [case for case in cases if str(case.get("path", "")).strip() != "A0"]
    return [a0_case] + non_a0_cases


def generate_test_cases(
    graph: dict[str, list[Transition]],
    *,
    root: str = "A1",
    max_depth: int = 200,
    emit_at_every_step: bool = False,
) -> list[dict[str, Any]]:
    cases: List[Dict[str, Any]] = []
    full_dfs_dst_owner: Dict[str, str] = {}

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

    def _append_case(
        case_conditions: List[str],
        steps: List[str],
        bot_responses_list: List[List[str]],
        action_code: str,
        path_nodes: List[str],
        source_rows: List[dict[str, str]],
    ) -> None:
        cases.append({
            "conditions": _render_conditions(case_conditions),
            "steps": steps,
            "bot_responses": bot_responses_list,
            "expected_action_code": action_code,
            "path": " -> ".join(path_nodes),
            "source_rows": list(source_rows),
        })

    # ------------------------------------------------------------------ #
    # Main DFS traversal (with optional stop‑after‑first‑terminal flag)
    # ------------------------------------------------------------------ #
    def dfs(
        node: str,
        steps: List[str],
        bot_responses_list: List[List[str]],
        case_conditions: List[str],
        path_nodes: List[str],
        source_rows: List[dict[str, str]],
        visited_edges: Set[Tuple[str, str, str, str]],
        consumed_chains: Set[Tuple[str, str]],
        *,
        stop_after_first: bool = False,
    ) -> bool:
        """
        Returns True if a terminal case was generated (useful when
        stop_after_first=True), otherwise returns False.
        """
        if len(steps) >= max_depth:
            return False

        transitions = graph.get(node, [])
        if not transitions:
            return False

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

        # For non-self-repeat transitions sharing the same dst:
        # 1) first occurrence gets full DFS,
        # 2) subsequent occurrences each stop after one terminal case.
        # Terminal transitions keep one case per (dst, condition, intent) because
        # there is no deeper DFS where stop_after_first can choose a path.
        # Self-repeat transitions (dst == node) are exempt from this rule.
        normal_sorted = sorted(
            normal_transitions,
            key=lambda tr: (tr.dst, tr.action_code, tr.condition, tr.intent),
        )
        dst_first: Dict[Tuple[str, str, str], Transition] = {}
        duplicates: List[Transition] = []
        self_repeats: List[Transition] = []
        for tr in normal_sorted:
            if tr.dst == node:
                self_repeats.append(tr)
            else:
                if _is_terminal_step(tr.dst):
                    group_key = (tr.dst, tr.condition, tr.intent)
                else:
                    group_key = (tr.dst, "", "")
                if group_key not in dst_first:
                    dst_first[group_key] = tr
                else:
                    duplicates.append(tr)

        normal_work: List[Tuple[Transition, bool, bool]] = []
        normal_work.extend((tr, stop_after_first, False) for tr in dst_first.values())
        normal_work.extend((tr, stop_after_first, False) for tr in self_repeats)
        normal_work.extend((tr, True, True) for tr in duplicates)

        for tr, effective_stop, is_duplicate_dst in normal_work:
            dst = tr.dst
            action_code = tr.action_code
            step_intent, bot_resp, step_condition = _make_single_step(tr)

            allowed, step_condition = _evaluate_step_condition(
                step_condition, {n.strip() for n in path_nodes}
            )
            if not allowed:
                continue

            edge_key = (node, dst, step_intent, step_condition)
            if edge_key in visited_edges:
                continue

            new_steps = steps + [step_intent]
            new_bot = bot_responses_list + [bot_resp]
            new_conditions = _append_condition(case_conditions, step_condition)
            new_path = path_nodes + [dst]
            new_source_rows = source_rows + [tr.source_row]
            new_visited = visited_edges | {edge_key}

            if _is_terminal_step(dst):
                if is_duplicate_dst:
                    continue
                _append_case(new_conditions, new_steps, new_bot, action_code, new_path, new_source_rows)
                if stop_after_first:
                    return True
                continue

            if emit_at_every_step:
                _append_case(new_conditions, new_steps, new_bot, action_code, new_path, new_source_rows)

            if dst in graph:
                cross_source_stop = effective_stop
                if dst != node:
                    owner = full_dfs_dst_owner.get(dst)
                    if owner is None:
                        if not effective_stop:
                            full_dfs_dst_owner[dst] = node
                    elif owner != node:
                        cross_source_stop = True
                found = dfs(
                    dst, new_steps, new_bot, new_conditions, new_path,
                    new_source_rows, new_visited, consumed_chains,
                    stop_after_first=cross_source_stop,
                )
                if stop_after_first and found:
                    return True

        # ── Process ordered repeat chains per base ────────────────────────
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

            found = _dfs_repeat_chain(
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
                stop_after_first=stop_after_first,
            )
            if stop_after_first and found:
                return True

        return False

    def _dfs_repeat_chain(
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
        *,
        stop_after_first: bool = False,
    ) -> bool:
        """Returns True if a terminal case was generated (only meaningful when stop_after_first=True)."""
        if chain_idx >= len(chain_levels):
            return False
        if len(steps) >= max_depth:
            return False

        _, grouped_transitions = chain_levels[chain_idx]
        is_last_in_chain = chain_idx == len(chain_levels) - 1

        for dst, action_code, step_condition, trans_list in grouped_transitions:
            step_intent, bot_resp, step_condition = _make_step(trans_list)

            new_path = path_nodes + [dst]
            allowed, step_condition = _evaluate_step_condition(
                step_condition, {n.strip() for n in new_path}
            )
            if not allowed:
                continue

            edge_key = (origin_node, dst, step_intent, step_condition)
            if edge_key in visited_edges:
                continue

            new_steps = steps + [step_intent]
            new_bot = bot_responses_list + [bot_resp]
            new_conditions = _append_condition(case_conditions, step_condition)
            new_source_rows = source_rows + [tr.source_row for tr in trans_list]
            new_visited = visited_edges | {edge_key}
            new_consumed = consumed_chains | chain_ids

            if is_last_in_chain:
                if _is_terminal_step(dst):
                    _append_case(new_conditions, new_steps, new_bot, action_code, new_path, new_source_rows)
                    if stop_after_first:
                        return True
                    continue
                if emit_at_every_step:
                    _append_case(new_conditions, new_steps, new_bot, action_code, new_path, new_source_rows)
                if dst in graph:
                    cross_source_stop = stop_after_first
                    if dst != origin_node:
                        owner = full_dfs_dst_owner.get(dst)
                        if owner is None:
                            if not stop_after_first:
                                full_dfs_dst_owner[dst] = origin_node
                        elif owner != origin_node:
                            cross_source_stop = True
                    found = dfs(
                        dst, new_steps, new_bot, new_conditions, new_path,
                        new_source_rows, new_visited, new_consumed,
                        stop_after_first=cross_source_stop,
                    )
                    if stop_after_first and found:
                        return True
            else:
                if emit_at_every_step:
                    _append_case(new_conditions, new_steps, new_bot, action_code, new_path, new_source_rows)
                found = _dfs_repeat_chain(
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
                    stop_after_first=stop_after_first,
                )
                if stop_after_first and found:
                    return True

        return False

    # Universal DFS entry point: per-dst transition grouping inside dfs() handles all nodes
    # (first transition per dst → full DFS; subsequent same-dst transitions → stop_after_first=True).
    # A cross-source owner map also limits later non-terminal transitions from
    # different source steps to the same dst to stop_after_first=True.
    # Self-repeat nodes and A0 are exempt (A0 is prepended separately below).
    dfs(root, [], [], [], [root], [], set(), set(), stop_after_first=False)

    # Remove duplicates
    seen = set()
    unique_cases = []
    for case in cases:
        key = (tuple(case["steps"]), case.get("conditions", ""), case["expected_action_code"])
        if key not in seen:
            seen.add(key)
            unique_cases.append(case)

    if root != "A0":
        return _prepend_single_a0_case(unique_cases, graph, multi_response=False)
    return unique_cases


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
    headers = ["tc_id", "conditions", "path", "expected_action_code",
               *[f"step_{i+1}" for i in range(max_len)],
               *[f"bot_response_{i+1}" for i in range(max_len)]]

    with out_path.open("w", encoding="utf-8", newline="") as f:
        w = csv.writer(f)
        w.writerow(headers)
        for i, tc in enumerate(cases):
            steps = tc["steps"]
            bot_responses_raw = tc.get("bot_responses", [])
            bot_responses = [
                " \\ ".join(resp_list) if resp_list else ""
                for resp_list in bot_responses_raw
            ]
            conditions = tc.get("conditions", "")
            action_code = tc.get("expected_action_code", "N/A")
            path = str(tc.get("path", ""))
            padded_steps = steps + [""] * (max_len - len(steps))
            padded_responses = bot_responses + [""] * (max_len - len(bot_responses))
            row = [f"TC{i+1:03d}", conditions, path, action_code] + padded_steps + padded_responses
            w.writerow(row)


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
