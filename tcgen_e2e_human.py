from __future__ import annotations

import argparse
import csv
import json
import re
from collections import defaultdict
from dataclasses import dataclass
from pathlib import Path
from typing import Any, DefaultDict, Iterable, Dict, List, Tuple, Set, Optional


@dataclass(frozen=True)
class Transition:
    src: str
    dst: str
    intent: str
    action_code: str
    bot_response: str


_REPEAT_RE = re.compile(r"(?i)\blần\s*(\d+)\b")


def _clean_str(v: Any) -> str:
    if v is None:
        return ""
    s = str(v).strip()
    return "" if s.lower() == "nan" else s


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


def _has_next_repeat(transitions: List[Transition], base: str, current_n: int) -> bool:
    """Check if next repeat (current_n + 1) exists in given transition list."""
    next_n = current_n + 1
    for tr in transitions:
        rep = _parse_repeat(tr.intent)
        if rep and rep[0] == base and rep[1] == next_n:
            return True
    return False


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
                "customer_intent": _clean_str(obj.get("customer_intent", "")),
                "bot_response": _clean_str(obj.get("bot_response", "")),
                "next_step": _clean_str(obj.get("next_step", "")),
                "action_code": _clean_str(obj.get("action_code", "")),
            }
        )
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
        intent = _clean_str(t.get("customer_intent", ""))
        if not intent:
            intent = "(empty intent)"
        action_code = _clean_str(t.get("action_code", ""))
        if not action_code:
            action_code = "N/A"
        bot_response = _clean_str(t.get("bot_response", ""))
        if not src:
            continue
        adjacency[src].append(
            Transition(src=src, dst=dst, intent=intent, action_code=action_code, bot_response=bot_response)
        )
    for src in adjacency:
        adjacency[src].sort(key=lambda tr: (tr.dst, tr.action_code, tr.intent))
    return dict(adjacency)


def generate_test_cases(
    graph: dict[str, list[Transition]],
    *,
    root: str = "A1",
    max_depth: int = 200,
) -> list[dict[str, Any]]:
    cases: List[Dict[str, Any]] = []

    def dfs(
        node: str,
        steps: List[str],
        bot_responses_list: List[List[str]],
        path_nodes: List[str],
        visited_edges: Set[Tuple[str, str, str]],  # (src, dst, intent_group_key)
    ) -> None:
        if len(steps) >= max_depth:
            return

        transitions = graph.get(node, [])
        if not transitions:
            return

        # ── Separate repeat vs non-repeat transitions ──────────────────────
        # Group repeat transitions by (base, n) — each lần N is its own group
        # Group non-repeat transitions by (dst, action_code)

        # Repeat groups: key = (base, n, dst, action_code)
        repeat_groups: Dict[Tuple[str, int, str, str], List[Transition]] = defaultdict(list)
        # Non-repeat groups: key = (dst, action_code)
        non_repeat_groups: Dict[Tuple[str, str], List[Transition]] = defaultdict(list)

        for tr in transitions:
            rep = _parse_repeat(tr.intent)
            if rep:
                base, n = rep
                repeat_groups[(base, n, tr.dst, tr.action_code)].append(tr)
            else:
                non_repeat_groups[(tr.dst, tr.action_code)].append(tr)

        # ── Process non-repeat groups ──────────────────────────────────────
        for (dst, action_code), trans_list in sorted(non_repeat_groups.items()):
            intents = sorted(set(tr.intent for tr in trans_list))
            step_intent = " \\ ".join(intents)
            bot_resp = list(dict.fromkeys(tr.bot_response for tr in trans_list if tr.bot_response))

            edge_key = (node, dst, step_intent)
            if edge_key in visited_edges:
                continue

            new_steps = steps + [step_intent]
            new_bot = bot_responses_list + [bot_resp]
            new_path = path_nodes + [dst]
            new_visited = visited_edges | {edge_key}

            if _is_terminal_step(dst):
                cases.append({
                    "steps": new_steps,
                    "bot_responses": new_bot,
                    "expected_action_code": action_code,
                    "path": " -> ".join(new_path),
                })
            elif dst in graph:
                dfs(dst, new_steps, new_bot, new_path, new_visited)

        # ── Process repeat groups — walk the full chain ────────────────────
        # Find all unique (base, dst_self, action_code) chains starting at lần 1
        # A chain is: lần 1 → lần 2 → ... → lần N (where lần N+1 doesn't exist)
        # Each lần in the chain must self-loop (dst == node) except possibly the last

        # Collect all bases that have lần 1 from this node
        chain_starts: Dict[Tuple[str, str, str], int] = {}  # (base, first_dst, first_ac) -> starting n
        for (base, n, dst, action_code) in repeat_groups:
            if n == 1:
                chain_starts[(base, dst, action_code)] = 1

        for (base, first_dst, first_ac) in sorted(chain_starts.keys()):
            # Walk chain: collect all lần in order
            chain: List[Tuple[int, str, str, List[Transition]]] = []  # (n, dst, ac, transitions)
            n = 1
            while True:
                # Find the group for (base, n, any_dst, any_ac)
                found = False
                for (b, num, dst, ac), trans_list in repeat_groups.items():
                    if b == base and num == n:
                        chain.append((n, dst, ac, trans_list))
                        found = True
                        break
                if not found:
                    break
                n += 1

            if not chain:
                continue

            # Now do DFS over the chain steps
            # We'll recurse step by step through the chain
            _dfs_repeat_chain(
                node=node,
                chain=chain,
                chain_idx=0,
                steps=steps,
                bot_responses_list=bot_responses_list,
                path_nodes=path_nodes,
                visited_edges=visited_edges,
                graph=graph,
                cases=cases,
                max_depth=max_depth,
                dfs_fn=dfs,
            )

    def _dfs_repeat_chain(
        node: str,
        chain: List[Tuple[int, str, str, List[Transition]]],
        chain_idx: int,
        steps: List[str],
        bot_responses_list: List[List[str]],
        path_nodes: List[str],
        visited_edges: Set[Tuple[str, str, str]],
        graph: dict,
        cases: List,
        max_depth: int,
        dfs_fn,
    ) -> None:
        """Recursively walk through a repeat chain, generating one TC per chain ending."""
        if chain_idx >= len(chain):
            return
        if len(steps) >= max_depth:
            return

        n, dst, action_code, trans_list = chain[chain_idx]
        intents = sorted(set(tr.intent for tr in trans_list))
        step_intent = " \\ ".join(intents)
        bot_resp = list(dict.fromkeys(tr.bot_response for tr in trans_list if tr.bot_response))

        edge_key = (node, dst, step_intent)
        if edge_key in visited_edges:
            return

        new_steps = steps + [step_intent]
        new_bot = bot_responses_list + [bot_resp]
        new_path = path_nodes + [dst]
        new_visited = visited_edges | {edge_key}

        is_last_in_chain = (chain_idx == len(chain) - 1)

        if _is_terminal_step(dst):
            # Terminal: emit TC and stop chain
            cases.append({
                "steps": new_steps,
                "bot_responses": new_bot,
                "expected_action_code": action_code,
                "path": " -> ".join(new_path),
            })
            return

        if is_last_in_chain:
            # Last repeat step, dst is non-terminal: continue with normal DFS from dst
            if dst in graph:
                dfs_fn(dst, new_steps, new_bot, new_path, new_visited)
            return

        # Not last: must continue chain (self-loop expected)
        # dst should be == node (self-loop), continue chain
        _dfs_repeat_chain(
            node=dst,  # next node (usually same as current for self-loop)
            chain=chain,
            chain_idx=chain_idx + 1,
            steps=new_steps,
            bot_responses_list=new_bot,
            path_nodes=new_path,
            visited_edges=new_visited,
            graph=graph,
            cases=cases,
            max_depth=max_depth,
            dfs_fn=dfs_fn,
        )

    dfs(root, [], [], [root], set())

    # Remove duplicates
    seen = set()
    unique_cases = []
    for case in cases:
        key = (tuple(case["steps"]), case["expected_action_code"])
        if key not in seen:
            seen.add(key)
            unique_cases.append(case)
    return unique_cases


def write_cases_json(cases: list[dict[str, Any]], out_path: Path) -> None:
    out_path.parent.mkdir(parents=True, exist_ok=True)
    payload = []
    for i, tc in enumerate(cases):
        bot_responses_str = [
            " \\ ".join(resp_list) if resp_list else ""
            for resp_list in tc.get("bot_responses", [])
        ]
        payload.append({
            "tc_id": f"TC{i+1:03d}",
            "steps": tc["steps"],
            "bot_responses": bot_responses_str,
            "expected_action_code": tc.get("expected_action_code", "N/A"),
            "path": tc.get("path", ""),
        })
    out_path.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")


def write_cases_csv(cases: list[dict[str, Any]], out_path: Path) -> None:
    out_path.parent.mkdir(parents=True, exist_ok=True)
    max_len = max((len(tc["steps"]) for tc in cases), default=0)
    headers = ["tc_id", "path", "expected_action_code",
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
            action_code = tc.get("expected_action_code", "N/A")
            path = str(tc.get("path", ""))
            padded_steps = steps + [""] * (max_len - len(steps))
            padded_responses = bot_responses + [""] * (max_len - len(bot_responses))
            row = [f"TC{i+1:03d}", path, action_code] + padded_steps + padded_responses
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
            bot_responses_str = [
                " \\ ".join(resp_list) if resp_list else ""
                for resp_list in tc.get("bot_responses", [])
            ]
            payload.append({
                "tc_id": f"TC{i+1:03d}",
                "steps": tc["steps"],
                "bot_responses": bot_responses_str,
                "expected_action_code": tc.get("expected_action_code", "N/A"),
                "path": tc.get("path", ""),
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