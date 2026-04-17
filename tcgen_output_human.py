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
    """
    Generate test cases by DFS over the transition graph.

    Key rules:
    - Non-repeat intents: group by (dst, action_code), emit one TC per group.
    - Repeat intents (lần N): each base+chain is kept SEPARATE from non-repeat
      and from other bases. Walk lần 1 → lần 2 → lần 3 in sequence.
    - visited_edges prevents re-traversing the same (src, dst, intent) edge.
    """
    cases: List[Dict[str, Any]] = []

    def _emit(steps, bot_responses_list, action_code, path_nodes):
        cases.append({
            "steps": list(steps),
            "bot_responses": [list(r) for r in bot_responses_list],
            "expected_action_code": action_code,
            "path": " -> ".join(path_nodes),
        })

    def _make_step(trans_list: List[Transition]) -> tuple[str, List[str]]:
        intents = sorted(set(tr.intent for tr in trans_list))
        step_intent = " \\ ".join(intents)
        bot_resp = list(dict.fromkeys(tr.bot_response for tr in trans_list if tr.bot_response))
        return step_intent, bot_resp

    def dfs(
        node: str,
        steps: List[str],
        bot_responses_list: List[List[str]],
        path_nodes: List[str],
        visited_edges: Set[Tuple[str, str, str]],  # (src, dst, intent)
        consumed_chains: Set[Tuple[str, str]],
    ) -> None:
        if len(steps) >= max_depth:
            return

        transitions = graph.get(node, [])
        if not transitions:
            return

        repeat_map: Dict[str, Dict[int, Dict[Tuple[str, str], List[Transition]]]] = defaultdict(
            lambda: defaultdict(lambda: defaultdict(list))
        )
        self_loop_group: Dict[Tuple[str, str], List[Transition]] = defaultdict(list)
        normal_group: Dict[Tuple[str, str], List[Transition]] = defaultdict(list)

        for tr in transitions:
            rep = _parse_repeat(tr.intent)
            if rep:
                base, n = rep
                repeat_map[base][n][(tr.dst, tr.action_code)].append(tr)
            elif tr.dst == node:
                self_loop_group[(tr.dst, tr.action_code)].append(tr)
            else:
                normal_group[(tr.dst, tr.action_code)].append(tr)

        # ── 1. Process normal groups ───────────────────────────────────────
        for (dst, action_code), trans_list in sorted(normal_group.items()):
            step_intent, bot_resp = _make_step(trans_list)
            edge_key = (node, dst, step_intent)

            if edge_key in visited_edges:
                continue

            new_steps = steps + [step_intent]
            new_bot = bot_responses_list + [bot_resp]
            new_path = path_nodes + [dst]
            new_visited = visited_edges | {edge_key}

            _emit(new_steps, new_bot, action_code, new_path)

            if not _is_terminal_step(dst) and dst in graph:
                dfs(dst, new_steps, new_bot, new_path, new_visited, consumed_chains)

        # ── 2. Process self-loop chains ────────────────────────────────────
        for (dst, action_code), trans_list in sorted(self_loop_group.items()):
            chain_id = (node, f"self:{dst}:{action_code}")
            if chain_id in consumed_chains:
                continue

            step_intent, bot_resp = _make_step(trans_list)
            edge_key = (node, dst, step_intent)
            if edge_key in visited_edges:
                continue

            new_steps = steps + [step_intent]
            new_bot = bot_responses_list + [bot_resp]
            new_path = path_nodes + [dst]
            new_visited = visited_edges | {edge_key}
            new_consumed = consumed_chains | {chain_id}

            _emit(new_steps, new_bot, action_code, new_path)
            if not _is_terminal_step(dst) and dst in graph:
                dfs(dst, new_steps, new_bot, new_path, new_visited, new_consumed)

        # ── 3. Process ordered repeat-chain clusters ───────────────────────
        repeat_clusters: Dict[
            Tuple[Tuple[Tuple[Tuple[str, str], ...], ...], int],
            Dict[str, Any],
        ] = defaultdict(lambda: defaultdict(lambda: defaultdict(list)))

        for base, n_map in repeat_map.items():
            if 1 not in n_map:
                continue

            signature_levels: List[Tuple[Tuple[str, str], ...]] = []
            expected_n = 1
            while expected_n in n_map:
                signature_levels.append(tuple(sorted(n_map[expected_n].keys())))
                expected_n += 1

            cluster_key = (tuple(signature_levels), len(signature_levels))
            cluster = repeat_clusters[cluster_key]
            cluster.setdefault("bases", set()).add(base)
            cluster.setdefault("levels", defaultdict(lambda: defaultdict(list)))
            for n, grouped in n_map.items():
                for group_key, trans_list in grouped.items():
                    cluster["levels"][n][group_key].extend(trans_list)

        for cluster_key, cluster_map in sorted(repeat_clusters.items(), key=lambda item: item[0]):
            base_ids = {("__repeat_base__", base) for base in cluster_map["bases"]}
            if base_ids & consumed_chains:
                continue

            chain_levels: List[Tuple[int, List[Tuple[str, str, List[Transition]]]]] = []
            for n in sorted(cluster_map["levels"].keys()):
                grouped = [
                    (dst, action_code, trans_list)
                    for (dst, action_code), trans_list in sorted(cluster_map["levels"][n].items())
                ]
                chain_levels.append((n, grouped))

            _walk_repeat_chain(
                origin_node=node,
                chain_ids=base_ids,
                chain_levels=chain_levels,
                chain_idx=0,
                steps=steps,
                bot_responses_list=bot_responses_list,
                path_nodes=path_nodes,
                visited_edges=visited_edges,
                consumed_chains=consumed_chains,
            )

    def _walk_repeat_chain(
        origin_node: str,
        chain_ids: Set[Tuple[str, str]],
        chain_levels: List[Tuple[int, List[Tuple[str, str, List[Transition]]]]],
        chain_idx: int,
        steps: List[str],
        bot_responses_list: List[List[str]],
        path_nodes: List[str],
        visited_edges: Set[Tuple[str, str, str]],
        consumed_chains: Set[Tuple[str, str]],
    ) -> None:
        if chain_idx >= len(chain_levels):
            return
        if len(steps) >= max_depth:
            return

        _, grouped_transitions = chain_levels[chain_idx]
        is_last = chain_idx == len(chain_levels) - 1

        for dst, action_code, trans_list in grouped_transitions:
            step_intent, bot_resp = _make_step(trans_list)
            edge_key = (origin_node, dst, step_intent)

            if edge_key in visited_edges:
                continue

            new_steps = steps + [step_intent]
            new_bot = bot_responses_list + [bot_resp]
            new_path = path_nodes + [dst]
            new_visited = visited_edges | {edge_key}
            new_consumed = consumed_chains | chain_ids

            if _is_terminal_step(dst):
                _emit(new_steps, new_bot, action_code, new_path)
                continue

            if is_last:
                _emit(new_steps, new_bot, action_code, new_path)
                if dst in graph:
                    dfs(dst, new_steps, new_bot, new_path, new_visited, new_consumed)
            else:
                _walk_repeat_chain(
                    origin_node=origin_node,
                    chain_ids=chain_ids,
                    chain_levels=chain_levels,
                    chain_idx=chain_idx + 1,
                    steps=new_steps,
                    bot_responses_list=new_bot,
                    path_nodes=new_path,
                    visited_edges=new_visited,
                    consumed_chains=new_consumed,
                )

    dfs(root, [], [], [root], set(), set())

    # Deduplicate
    seen: Set[Tuple] = set()
    unique: List[Dict[str, Any]] = []
    for case in cases:
        key = (tuple(case["steps"]), case["expected_action_code"])
        if key not in seen:
            seen.add(key)
            unique.append(case)
    return unique


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
    headers = [
        "tc_id", "path", "expected_action_code",
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
            action_code = tc.get("expected_action_code", "N/A")
            path = str(tc.get("path", ""))
            padded_steps = steps + [""] * (max_len - len(steps))
            padded_resp = bot_responses + [""] * (max_len - len(bot_responses))
            w.writerow([f"TC{i+1:03d}", path, action_code] + padded_steps + padded_resp)


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
                " \\ ".join(r) if r else ""
                for r in tc.get("bot_responses", [])
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
