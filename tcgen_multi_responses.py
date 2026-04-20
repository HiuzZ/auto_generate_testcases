from __future__ import annotations

import argparse
import csv
import json
import re
from collections import defaultdict
from dataclasses import dataclass
from pathlib import Path
from typing import Any, DefaultDict, Dict, Iterable, List, Tuple


@dataclass(frozen=True)
class Transition:
    src: str
    dst: str
    condition: str
    intent: str
    action_code: str
    responses: Tuple[str, ...]


@dataclass(frozen=True)
class GroupedTransition:
    src: str
    dst: str
    condition: str
    intents: Tuple[str, ...]
    action_code: str
    responses: Tuple[str, ...]


def _clean_str(v: Any) -> str:
    if v is None:
        return ""
    s = str(v).strip()
    return "" if s.lower() == "nan" else s


def _is_terminal_step(step_no: str) -> bool:
    s = step_no.strip().lower()
    return s in {"end", "stop", "__end__", "__terminal__"}


_REPEAT_INTENT_RE = re.compile(r"^(?P<base>.+?)\s+lần\s+(?P<count>\d+)\s*$", re.IGNORECASE)


def _normalize_spaces(text: str) -> str:
    return " ".join(text.strip().split())


def _repeat_prerequisite(intent: str) -> tuple[str, int] | None:
    match = _REPEAT_INTENT_RE.match(_normalize_spaces(intent))
    if not match:
        return None
    base = _normalize_spaces(match.group("base")).lower()
    count = int(match.group("count"))
    return (base, count)


def _can_use_transition(intents: Tuple[str, ...], seen_repeat_intents: set[tuple[str, int]]) -> bool:
    for intent in intents:
        repeat_info = _repeat_prerequisite(intent)
        if not repeat_info:
            continue
        base, count = repeat_info
        if count <= 1:
            continue
        if (base, count - 1) not in seen_repeat_intents:
            return False
    return True


def load_transitions_json(path: Path) -> list[dict[str, str]]:
    data = json.loads(path.read_text(encoding="utf-8"))
    if not isinstance(data, list):
        raise ValueError("Input must be a JSON list of transition objects.")
    rows: list[dict[str, str]] = []
    for i, obj in enumerate(data):
        if not isinstance(obj, dict):
            raise ValueError(f"Item at index {i} is not an object.")
        rows.append(
            {
                "step_no": _clean_str(obj.get("step_no", "")),
                "step_name": _clean_str(obj.get("step_name", "")),
                "conditions": _clean_str(obj.get("conditions", "")),
                "customer_intent": _clean_str(obj.get("customer_intent", "")),
                "bot_response": _clean_str(obj.get("bot_response", "")),
                "bot_response_2": _clean_str(obj.get("bot_response_2", "")),
                "bot_response_3": _clean_str(obj.get("bot_response_3", "")),
                "bot_response_4": _clean_str(obj.get("bot_response_4", "")),
                "bot_response_5": _clean_str(obj.get("bot_response_5", "")),
                "next_step": _clean_str(obj.get("next_step", "")),
                "action_code": _clean_str(obj.get("action_code", "")),
            }
        )
    return rows


def build_graph(
    transitions: Iterable[dict[str, str]],
) -> dict[str, list[GroupedTransition]]:
    adjacency: DefaultDict[str, list[Transition]] = defaultdict(list)
    for t in transitions:
        src = _clean_str(t.get("step_no", ""))
        dst = _clean_str(t.get("next_step", "")) or "End"
        if not src:
            continue

        intent = _clean_str(t.get("customer_intent", "")) or "(empty intent)"
        condition = _clean_str(t.get("conditions", ""))
        action_code = _clean_str(t.get("action_code", "")) or "N/A"
        responses = tuple(
            resp
            for resp in [
                _clean_str(t.get("bot_response", "")),
                _clean_str(t.get("bot_response_2", "")),
                _clean_str(t.get("bot_response_3", "")),
                _clean_str(t.get("bot_response_4", "")),
                _clean_str(t.get("bot_response_5", "")),
            ]
            if resp
        )
        adjacency[src].append(
            Transition(
                src=src,
                dst=dst,
                condition=condition,
                intent=intent,
                action_code=action_code,
                responses=responses,
            )
        )

    grouped_graph: dict[str, list[GroupedTransition]] = {}
    for src, src_transitions in adjacency.items():
        grouped: DefaultDict[Tuple[str, str, str, Tuple[str, ...]], list[Transition]] = defaultdict(list)
        for tr in src_transitions:
            grouped[(tr.dst, tr.action_code, tr.condition, tr.responses)].append(tr)

        grouped_graph[src] = sorted(
            [
                GroupedTransition(
                    src=src,
                    dst=dst,
                    condition=condition,
                    intents=tuple(sorted(set(t.intent for t in trans_list))),
                    action_code=action_code,
                    responses=responses,
                )
                for (dst, action_code, condition, responses), trans_list in grouped.items()
            ],
            key=lambda tr: (tr.dst, tr.action_code, tr.condition, tr.intents),
        )
    return grouped_graph


def generate_test_cases(
    graph: dict[str, list[GroupedTransition]],
    *,
    root: str = "A1",
    max_depth: int = 200,
) -> list[dict[str, Any]]:
    def _append_condition(case_conditions: List[str], condition: str) -> List[str]:
        if condition in case_conditions:
            return list(case_conditions)
        return case_conditions + [condition]

    def _render_conditions(case_conditions: List[str]) -> str:
        non_empty = [cond for cond in case_conditions if cond]
        return " \\ ".join(non_empty) if non_empty else ""

    repeat_index: Dict[str, Dict[Tuple[str, int], GroupedTransition]] = defaultdict(dict)
    for src, transitions in graph.items():
        for tr in transitions:
            if len(tr.intents) != 1:
                continue
            repeat_info = _repeat_prerequisite(tr.intents[0])
            if repeat_info:
                repeat_index[src][repeat_info] = tr

    queue: List[dict[str, Any]] = [
        {
            "node": root,
            "steps": [],
            "bot_responses": [],
            "conditions": [],
            "path_nodes": [root],
            "visited_nodes": {root},
            "visited_edges": set(),
            "seen_repeat_intents": set(),
        }
    ]

    cases: List[dict[str, Any]] = []
    seen_targets: set[Tuple[str, str, str, str, Tuple[str, ...], Tuple[str, ...]]] = set()

    while queue:
        state = queue.pop(0)
        if len(state["steps"]) >= max_depth:
            continue

        for tr in graph.get(state["node"], []):
            # Allow self-loop transitions (e.g. A1 -> A1) but prevent revisiting other nodes.
            if tr.dst != state["node"] and tr.dst in state["visited_nodes"]:
                continue

            curr_seen_repeat_intents = set(state["seen_repeat_intents"])
            chain_steps: List[str] = []
            chain_bot_responses: List[str] = []
            chain_conditions = list(state["conditions"])
            chain_path_nodes = list(state["path_nodes"])
            chain_visited_edges = set(state["visited_edges"])

            for intent in tr.intents:
                repeat_info = _repeat_prerequisite(intent)
                if not repeat_info:
                    continue
                base, n = repeat_info
                for k in range(1, n):
                    if (base, k) in curr_seen_repeat_intents:
                        continue
                    pre_tr = repeat_index.get(state["node"], {}).get((base, k))
                    if not pre_tr:
                        break
                    pre_step_text = " \\ ".join(pre_tr.intents)
                    chain_steps.append(pre_step_text)
                    chain_bot_responses.append(pre_tr.responses[0] if pre_tr.responses else "")
                    chain_conditions = _append_condition(chain_conditions, pre_tr.condition)
                    curr_seen_repeat_intents.add((base, k))

            if not _can_use_transition(tr.intents, curr_seen_repeat_intents):
                continue

            step_text = " \\ ".join(tr.intents)
            edge_key = (tr.src, tr.dst, step_text, tr.condition)
            if edge_key in chain_visited_edges:
                continue
            new_conditions = _append_condition(chain_conditions, tr.condition)
            new_steps = state["steps"] + chain_steps + [step_text]
            new_path = chain_path_nodes + [tr.dst]
            is_multi = len(tr.responses) > 1
            new_visited_edges = chain_visited_edges | {edge_key}
            new_seen_repeat_intents = set(curr_seen_repeat_intents)
            for intent in tr.intents:
                repeat_info = _repeat_prerequisite(intent)
                if repeat_info:
                    new_seen_repeat_intents.add(repeat_info)

            if is_multi:
                target_key = (
                    tr.src, tr.dst, step_text, tr.condition, tr.responses, tuple(chain_steps)
                )
                if target_key not in seen_targets:
                    seen_targets.add(target_key)
                    prior_bot = state["bot_responses"] + chain_bot_responses
                    for response in tr.responses:
                        cases.append(
                            {
                                "conditions": _render_conditions(new_conditions),
                                "steps": new_steps,
                                "bot_responses": prior_bot + [response],
                                "expected_action_code": tr.action_code,
                                "path": " -> ".join(new_path),
                                "highlight_last_step": True,
                            }
                        )

            # IMPORTANT: do NOT traverse self-loops further (prevents combinatorial explosion).
            if tr.dst != state["node"] and (not _is_terminal_step(tr.dst)) and tr.dst in graph:
                queue.append(
                    {
                        "node": tr.dst,
                        "steps": new_steps,
                        "bot_responses": state["bot_responses"] + chain_bot_responses + [tr.responses[0] if tr.responses else ""],
                        "conditions": new_conditions,
                        "path_nodes": new_path,
                        "visited_nodes": state["visited_nodes"] | ({tr.dst} if tr.dst != state["node"] else set()),
                        "visited_edges": new_visited_edges,
                        "seen_repeat_intents": new_seen_repeat_intents,
                    }
                )

    seen: set[Tuple[Any, ...]] = set()
    unique_cases: List[dict[str, Any]] = []
    for case in cases:
        key = (
            tuple(case["steps"]),
            tuple(case["bot_responses"]),
            case.get("conditions", ""),
            case["expected_action_code"],
            case["path"],
        )
        if key not in seen:
            seen.add(key)
            unique_cases.append(case)

    # De-duplicate redundant cases by last step: if a shorter testcase already covers the same
    # last step + last bot response + expected action code, drop the longer one.
    unique_cases.sort(key=lambda c: (len(c.get("steps", [])), c.get("tc_id", "")))
    kept: List[dict[str, Any]] = []
    covered_last: set[Tuple[str, str, str]] = set()
    for case in unique_cases:
        steps = case.get("steps", [])
        bots = case.get("bot_responses", [])
        last_step = steps[-1] if steps else ""
        last_bot = bots[-1] if bots else ""
        action = str(case.get("expected_action_code", ""))
        last_key = (last_step, last_bot, action)
        if last_key in covered_last:
            continue
        covered_last.add(last_key)
        kept.append(case)
    return kept


def write_cases_json(cases: list[dict[str, Any]], out_path: Path) -> None:
    out_path.parent.mkdir(parents=True, exist_ok=True)
    payload = []
    for i, tc in enumerate(cases):
        payload.append(
            {
                "tc_id": f"TC{i+1:03d}",
                "conditions": tc.get("conditions", ""),
                "steps": tc["steps"],
                "bot_responses": tc.get("bot_responses", []),
                "expected_action_code": tc.get("expected_action_code", "N/A"),
                "path": tc.get("path", ""),
                "highlight_last_step": bool(tc.get("highlight_last_step")),
            }
        )
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
            bot_responses = [str(r) for r in tc.get("bot_responses", [])]
            padded_steps = steps + [""] * (max_len - len(steps))
            padded_responses = bot_responses + [""] * (max_len - len(bot_responses))
            w.writerow(
                [
                    f"TC{i+1:03d}",
                    tc.get("conditions", ""),
                    tc.get("path", ""),
                    tc.get("expected_action_code", "N/A"),
                    *padded_steps,
                    *padded_responses,
                ]
            )


def main() -> int:
    parser = argparse.ArgumentParser(description="Generate shortest-path multi-response test cases.")
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
            payload.append(
                {
                    "tc_id": f"TC{i+1:03d}",
                    "conditions": tc.get("conditions", ""),
                    "steps": tc["steps"],
                    "bot_responses": tc.get("bot_responses", []),
                    "expected_action_code": tc.get("expected_action_code", "N/A"),
                    "path": tc.get("path", ""),
                    "highlight_last_step": bool(tc.get("highlight_last_step")),
                }
            )
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
