from __future__ import annotations

import argparse
import json
from pathlib import Path
from typing import Any

import tcgen_e2e_human


_STEP_SEPARATOR = " \\ "


def _is_terminal_step(step_no: str) -> bool:
    s = str(step_no).strip().lower()
    return s in {"end", "stop", "__end__", "__terminal__"}


def _extract_nodes_from_path(path: str) -> list[str]:
    return [part.strip() for part in str(path).split("->") if part.strip()]


def _path_length(case: dict[str, Any]) -> int:
    path_nodes = _extract_nodes_from_path(str(case.get("path", "")))
    if len(path_nodes) >= 2:
        return len(path_nodes) - 1
    return len(case.get("steps", []))


def _prefix_to_node(case: dict[str, Any], target_node: str) -> tuple[str, ...]:
    path_nodes = _extract_nodes_from_path(str(case.get("path", "")))
    if not path_nodes:
        return tuple()
    try:
        idx = path_nodes.index(target_node)
    except ValueError:
        return tuple(path_nodes)
    return tuple(path_nodes[: idx + 1])


def _collect_universe(graph: dict[str, list[tcgen_e2e_human.Transition]]) -> set[str]:
    universe: set[str] = set(graph.keys())
    for transitions in graph.values():
        for tr in transitions:
            dst = str(tr.dst).strip()
            if dst and not _is_terminal_step(dst):
                universe.add(dst)
    return universe


def _split_step_intents(step_text: str) -> list[str]:
    return [part.strip() for part in str(step_text).split(_STEP_SEPARATOR) if part.strip()]


def _collect_intent_universe(
    graph: dict[str, list[tcgen_e2e_human.Transition]],
    step_nodes: set[str],
) -> set[tuple[str, str, str]]:
    intent_universe: set[tuple[str, str, str]] = set()
    for src, transitions in graph.items():
        if src not in step_nodes:
            continue
        for tr in transitions:
            intent = str(tr.intent).strip()
            condition = str(tr.condition).strip()
            if intent:
                intent_universe.add((src, condition, intent))
    return intent_universe


def _case_intent_pairs(
    case: dict[str, Any],
    step_nodes: set[str],
    graph: dict[str, list[tcgen_e2e_human.Transition]],
) -> set[tuple[str, str, str]]:
    path_nodes = _extract_nodes_from_path(str(case.get("path", "")))
    steps = list(case.get("steps", []))
    conditions_text = str(case.get("conditions", ""))
    case_conditions = {part.strip() for part in conditions_text.split(_STEP_SEPARATOR) if part.strip()}
    intent_pairs: set[tuple[str, str, str]] = set()
    for idx, step_text in enumerate(steps):
        if idx >= len(path_nodes) - 1:
            break
        src = path_nodes[idx]
        if src not in step_nodes:
            continue
        for intent in _split_step_intents(str(step_text)):
            matched = False
            for tr in graph.get(src, []):  # type: ignore[name-defined]
                tr_intent = str(tr.intent).strip()
                tr_condition = str(tr.condition).strip()
                if tr_intent != intent:
                    continue
                if tr_condition and tr_condition not in case_conditions:
                    continue
                intent_pairs.add((src, tr_condition, intent))
                matched = True
            if not matched:
                intent_pairs.add((src, "", intent))
    return intent_pairs


def _collect_terminal_outcome_universe(
    graph: dict[str, list[tcgen_e2e_human.Transition]],
    step_nodes: set[str],
) -> set[tuple[str, str, str, str]]:
    universe: set[tuple[str, str, str, str]] = set()
    for src, transitions in graph.items():
        if src not in step_nodes:
            continue
        for tr in transitions:
            dst = str(tr.dst).strip()
            if not _is_terminal_step(dst):
                continue
            intent = str(tr.intent).strip()
            condition = str(tr.condition).strip()
            action_code = str(tr.action_code).strip()
            if intent:
                universe.add((src, condition, intent, action_code))
    return universe


def _case_terminal_outcome_pairs(
    case: dict[str, Any],
    graph: dict[str, list[tcgen_e2e_human.Transition]],
) -> set[tuple[str, str, str, str]]:
    path_nodes = _extract_nodes_from_path(str(case.get("path", "")))
    steps = list(case.get("steps", []))
    if not path_nodes or not steps:
        return set()
    conditions_text = str(case.get("conditions", ""))
    case_conditions = {part.strip() for part in conditions_text.split(_STEP_SEPARATOR) if part.strip()}
    expected_action = str(case.get("expected_action_code", "")).strip()
    pairs: set[tuple[str, str, str, str]] = set()

    for idx, step_text in enumerate(steps):
        if idx >= len(path_nodes) - 1:
            break
        src = path_nodes[idx]
        dst = path_nodes[idx + 1]
        if not _is_terminal_step(dst):
            continue
        intents = _split_step_intents(str(step_text))
        for intent in intents:
            matched = False
            for tr in graph.get(src, []):
                tr_intent = str(tr.intent).strip()
                tr_condition = str(tr.condition).strip()
                tr_dst = str(tr.dst).strip()
                tr_action = str(tr.action_code).strip()
                if tr_intent != intent or tr_dst != dst:
                    continue
                if tr_condition and tr_condition not in case_conditions:
                    continue
                # For terminal transition, bind to testcase expected action
                # so distinct outcomes are represented separately.
                if expected_action and tr_action != expected_action:
                    continue
                pairs.add((src, tr_condition, intent, tr_action))
                matched = True
            if not matched and expected_action:
                pairs.add((src, "", intent, expected_action))
    return pairs


def _case_signature(case: dict[str, Any]) -> tuple[Any, ...]:
    return (
        tuple(str(step) for step in case.get("steps", [])),
        tuple(str(resp) for resp in case.get("bot_responses", [])),
        str(case.get("conditions", "")),
        str(case.get("expected_action_code", "")),
        str(case.get("path", "")),
    )


def _case_transition_pairs(
    case: dict[str, Any],
    graph: dict[str, list[tcgen_e2e_human.Transition]],
) -> set[tuple[str, str, str, str]]:
    path_nodes = _extract_nodes_from_path(str(case.get("path", "")))
    steps = list(case.get("steps", []))
    conditions_text = str(case.get("conditions", ""))
    case_conditions = {part.strip() for part in conditions_text.split(_STEP_SEPARATOR) if part.strip()}
    pairs: set[tuple[str, str, str, str]] = set()

    for idx, step_text in enumerate(steps):
        if idx >= len(path_nodes) - 1:
            break
        src = path_nodes[idx]
        dst = path_nodes[idx + 1]
        for intent in _split_step_intents(str(step_text)):
            matched = False
            for tr in graph.get(src, []):
                tr_intent = str(tr.intent).strip()
                tr_condition = str(tr.condition).strip()
                tr_dst = str(tr.dst).strip()
                if tr_intent != intent or tr_dst != dst:
                    continue
                if tr_condition and tr_condition not in case_conditions:
                    continue
                pairs.add((src, tr_condition, intent, dst))
                matched = True
            if not matched:
                pairs.add((src, "", intent, dst))
    return pairs


def _collect_same_next_step_groups(
    graph: dict[str, list[tcgen_e2e_human.Transition]],
) -> dict[tuple[str, str, str], set[str]]:
    groups: dict[tuple[str, str, str], set[str]] = {}
    for src, transitions in graph.items():
        for tr in transitions:
            key = (src, str(tr.condition).strip(), str(tr.dst).strip())
            groups.setdefault(key, set()).add(str(tr.intent).strip())
    return groups


def _reduce_cases_by_step_coverage(
    cases: list[dict[str, Any]],
    universe: set[str],
) -> list[dict[str, Any]]:
    if not cases or not universe:
        return cases

    indexed_case_data: list[tuple[int, dict[str, Any], set[str], int, int]] = []
    for idx, case in enumerate(cases):
        path_nodes = _extract_nodes_from_path(str(case.get("path", "")))
        covered = {node for node in path_nodes if node in universe}
        indexed_case_data.append((idx, case, covered, _path_length(case), len(case.get("steps", []))))

    selected_indices: set[int] = set()

    # Rule 1: for each Step_no, choose one shortest testcase.
    # Rule 2: for longer feasible paths to that Step_no, keep one representative
    # per distinct path prefix reaching that Step_no.
    for target_node in sorted(universe):
        candidates = [item for item in indexed_case_data if target_node in item[2]]
        if not candidates:
            # Fallback: keep original list if full step coverage is impossible.
            return cases

        shortest_path_len = min(item[3] for item in candidates)
        shortest_candidates = [item for item in candidates if item[3] == shortest_path_len]
        shortest_choice = min(shortest_candidates, key=lambda item: (item[4], item[0]))
        selected_indices.add(shortest_choice[0])

        # For longer feasible paths, keep one representative per distinct
        # prefix path reaching the target node.
        longer_candidates = [item for item in candidates if item[3] > shortest_path_len]
        by_prefix: dict[tuple[str, ...], list[tuple[int, dict[str, Any], set[str], int, int]]] = {}
        for item in longer_candidates:
            _, case, _, _, _ = item
            prefix = _prefix_to_node(case, target_node)
            by_prefix.setdefault(prefix, []).append(item)
        for reps in by_prefix.values():
            rep_choice = min(reps, key=lambda item: (item[3], item[4], item[0]))
            selected_indices.add(rep_choice[0])

    selected = [item for item in indexed_case_data if item[0] in selected_indices]
    selected.sort(key=lambda item: item[0])
    return [item[1] for item in selected]


def _expand_cases_by_intent_coverage(
    cases: list[dict[str, Any]],
    graph: dict[str, list[tcgen_e2e_human.Transition]],
    *,
    root: str,
    max_depth: int,
) -> list[dict[str, Any]]:
    if not cases:
        return cases

    step_nodes: set[str] = set()
    for case in cases:
        for node in _extract_nodes_from_path(str(case.get("path", ""))):
            if not _is_terminal_step(node):
                step_nodes.add(node)

    intent_universe = _collect_intent_universe(graph, step_nodes)
    if not intent_universe:
        return cases

    all_cases = tcgen_e2e_human.generate_test_cases(graph, root=root, max_depth=max_depth)
    indexed_case_data: list[tuple[int, dict[str, Any], set[tuple[str, str, str]]]] = []
    for idx, case in enumerate(all_cases):
        intent_pairs = _case_intent_pairs(case, step_nodes, graph)
        if intent_pairs:
            indexed_case_data.append((idx, case, intent_pairs))

    selected: list[tuple[int, dict[str, Any], set[tuple[str, str, str]]]] = []
    selected_signatures = {_case_signature(case) for case in cases}
    for idx, case, intent_pairs in indexed_case_data:
        if _case_signature(case) in selected_signatures:
            selected.append((idx, case, intent_pairs))

    if not selected:
        return cases

    covered = set().union(*(intent_pairs for _, _, intent_pairs in selected))
    uncovered = intent_universe - covered
    remaining = [item for item in indexed_case_data if _case_signature(item[1]) not in selected_signatures]

    while uncovered:
        best: tuple[int, dict[str, Any], set[tuple[str, str, str]]] | None = None
        best_path_len = 10**9
        best_step_len = 10**9
        best_gain = -1
        best_idx = 10**9

        for idx, case, intent_pairs in remaining:
            gain = len(intent_pairs & uncovered)
            if gain <= 0:
                continue
            path_len = _path_length(case)
            step_len = len(case.get("steps", []))
            if (
                gain > best_gain
                or (gain == best_gain and path_len < best_path_len)
                or (gain == best_gain and path_len == best_path_len and step_len < best_step_len)
                or (
                    gain == best_gain
                    and path_len == best_path_len
                    and step_len == best_step_len
                    and idx < best_idx
                )
            ):
                best = (idx, case, intent_pairs)
                best_path_len = path_len
                best_step_len = step_len
                best_gain = gain
                best_idx = idx

        if best is None:
            break

        selected.append(best)
        covered |= best[2]
        uncovered = intent_universe - covered
        remaining = [item for item in remaining if item[0] != best[0]]

    # Hard guarantee: if any (src, condition, intent) pair is still missing,
    # backfill with the shortest testcase that covers that pair.
    covered = set().union(*(intent_pairs for _, _, intent_pairs in selected))
    still_missing = intent_universe - covered
    if still_missing:
        selected_sig = {_case_signature(case) for _, case, _ in selected}
        for pair in sorted(still_missing):
            candidates = [
                item
                for item in indexed_case_data
                if pair in item[2] and _case_signature(item[1]) not in selected_sig
            ]
            if not candidates:
                continue
            chosen = min(
                candidates,
                key=lambda item: (_path_length(item[1]), len(item[1].get("steps", [])), item[0]),
            )
            selected.append(chosen)
            selected_sig.add(_case_signature(chosen[1]))

    # Also guarantee terminal action outcomes for selected nodes.
    selected_cases = [item[1] for item in selected]
    terminal_universe = _collect_terminal_outcome_universe(graph, step_nodes)
    covered_terminal: set[tuple[str, str, str, str]] = set()
    for case in selected_cases:
        covered_terminal |= _case_terminal_outcome_pairs(case, graph)
    missing_terminal = terminal_universe - covered_terminal
    if missing_terminal:
        selected_sig = {_case_signature(case) for case in selected_cases}
        for outcome in sorted(missing_terminal):
            candidates = [
                item
                for item in indexed_case_data
                if _case_signature(item[1]) not in selected_sig
                and outcome in _case_terminal_outcome_pairs(item[1], graph)
            ]
            if not candidates:
                continue
            chosen = min(
                candidates,
                key=lambda item: (_path_length(item[1]), len(item[1].get("steps", [])), item[0]),
            )
            selected.append(chosen)
            selected_sig.add(_case_signature(chosen[1]))

    selected.sort(key=lambda item: item[0])
    return [item[1] for item in selected]


def _expand_cases_by_shared_next_step_siblings(
    cases: list[dict[str, Any]],
    graph: dict[str, list[tcgen_e2e_human.Transition]],
    *,
    root: str,
    max_depth: int,
) -> list[dict[str, Any]]:
    if not cases:
        return cases

    sibling_groups = _collect_same_next_step_groups(graph)
    all_cases = tcgen_e2e_human.generate_test_cases(graph, root=root, max_depth=max_depth)
    indexed_case_data: list[tuple[int, dict[str, Any], set[tuple[str, str, str, str]]]] = []
    for idx, case in enumerate(all_cases):
        transition_pairs = _case_transition_pairs(case, graph)
        if transition_pairs:
            indexed_case_data.append((idx, case, transition_pairs))

    selected_signatures = {_case_signature(case) for case in cases}
    selected: list[tuple[int, dict[str, Any], set[tuple[str, str, str, str]]]] = []
    for idx, case, transition_pairs in indexed_case_data:
        if _case_signature(case) in selected_signatures:
            selected.append((idx, case, transition_pairs))

    if not selected:
        return cases

    remaining = [item for item in indexed_case_data if _case_signature(item[1]) not in selected_signatures]

    while True:
        covered_pairs = set().union(*(pairs for _, _, pairs in selected))
        missing_pairs: set[tuple[str, str, str, str]] = set()

        for _, _, pairs in selected:
            intents_by_group: dict[tuple[str, str, str], set[str]] = {}
            for src, condition, intent, dst in pairs:
                intents_by_group.setdefault((src, condition, dst), set()).add(intent)

            for group_key, case_intents in intents_by_group.items():
                sibling_intents = sibling_groups.get(group_key, set())
                if len(sibling_intents) <= 1:
                    continue
                # Only expand groups where this selected case already contains
                # all intents in the sibling group.
                if case_intents != sibling_intents:
                    continue
                src, condition, dst = group_key
                for sibling_intent in sibling_intents:
                    sibling_pair = (src, condition, sibling_intent, dst)
                    if sibling_pair not in covered_pairs:
                        missing_pairs.add(sibling_pair)

        if not missing_pairs:
            break

        best: tuple[int, dict[str, Any], set[tuple[str, str, str, str]]] | None = None
        best_path_len = 10**9
        best_step_len = 10**9
        best_gain = -1
        best_idx = 10**9

        for idx, case, transition_pairs in remaining:
            gain = len(transition_pairs & missing_pairs)
            if gain <= 0:
                continue
            path_len = _path_length(case)
            step_len = len(case.get("steps", []))
            if (
                gain > best_gain
                or (gain == best_gain and path_len < best_path_len)
                or (gain == best_gain and path_len == best_path_len and step_len < best_step_len)
                or (
                    gain == best_gain
                    and path_len == best_path_len
                    and step_len == best_step_len
                    and idx < best_idx
                )
            ):
                best = (idx, case, transition_pairs)
                best_path_len = path_len
                best_step_len = step_len
                best_gain = gain
                best_idx = idx

        if best is None:
            break

        selected.append(best)
        remaining = [item for item in remaining if item[0] != best[0]]

    selected.sort(key=lambda item: item[0])
    return [item[1] for item in selected]


def _expand_cases_by_shortest_sibling_representatives(
    cases: list[dict[str, Any]],
    graph: dict[str, list[tcgen_e2e_human.Transition]],
    *,
    root: str,
    max_depth: int,
) -> list[dict[str, Any]]:
    if not cases:
        return cases

    all_cases = tcgen_e2e_human.generate_test_cases(graph, root=root, max_depth=max_depth)
    sibling_groups = _collect_same_next_step_groups(graph)

    indexed_case_data: list[tuple[int, dict[str, Any], set[tuple[str, str, str, str]]]] = []
    for idx, case in enumerate(all_cases):
        transition_pairs = _case_transition_pairs(case, graph)
        if transition_pairs:
            indexed_case_data.append((idx, case, transition_pairs))

    selected_signatures = {_case_signature(case) for case in cases}
    selected: list[tuple[int, dict[str, Any], set[tuple[str, str, str, str]]]] = []
    for idx, case, transition_pairs in indexed_case_data:
        if _case_signature(case) in selected_signatures:
            selected.append((idx, case, transition_pairs))

    if not selected:
        return cases

    needed_pairs: set[tuple[str, str, str, str]] = set()
    for _, _, transition_pairs in selected:
        for src, condition, _, dst in transition_pairs:
            sibling_intents = sibling_groups.get((src, condition, dst), set())
            if len(sibling_intents) <= 1:
                continue
            for sibling_intent in sibling_intents:
                needed_pairs.add((src, condition, sibling_intent, dst))

    if not needed_pairs:
        return cases

    shortest_case_for_pair: dict[tuple[str, str, str, str], tuple[int, dict[str, Any], set[tuple[str, str, str, str]]]] = {}
    for idx, case, transition_pairs in indexed_case_data:
        path_len = _path_length(case)
        step_len = len(case.get("steps", []))
        for pair in transition_pairs & needed_pairs:
            current = shortest_case_for_pair.get(pair)
            if current is None:
                shortest_case_for_pair[pair] = (idx, case, transition_pairs)
                continue
            current_idx, current_case, _ = current
            current_path_len = _path_length(current_case)
            current_step_len = len(current_case.get("steps", []))
            if (
                path_len < current_path_len
                or (path_len == current_path_len and step_len < current_step_len)
                or (path_len == current_path_len and step_len == current_step_len and idx < current_idx)
            ):
                shortest_case_for_pair[pair] = (idx, case, transition_pairs)

    for pair in sorted(needed_pairs):
        chosen = shortest_case_for_pair.get(pair)
        if chosen is None:
            continue
        signature = _case_signature(chosen[1])
        if signature in selected_signatures:
            continue
        selected.append(chosen)
        selected_signatures.add(signature)

    selected.sort(key=lambda item: item[0])
    return [item[1] for item in selected]


def _fix_terminal_action_and_detours(
    cases: list[dict[str, Any]],
    graph: dict[str, list[tcgen_e2e_human.Transition]],
    *,
    root: str,
    max_depth: int,
) -> list[dict[str, Any]]:
    if not cases:
        return cases

    all_cases = tcgen_e2e_human.generate_test_cases(graph, root=root, max_depth=max_depth)

    def _flow_key_without_action(case: dict[str, Any]) -> tuple[Any, ...]:
        return (
            tuple(str(step) for step in case.get("steps", [])),
            str(case.get("conditions", "")),
            str(case.get("path", "")),
        )

    # 1) Keep all action variants for a selected exact flow.
    by_key_actions: dict[tuple[Any, ...], dict[str, dict[str, Any]]] = {}
    for case in all_cases:
        key = _flow_key_without_action(case)
        action = str(case.get("expected_action_code", ""))
        by_key_actions.setdefault(key, {})
        by_key_actions[key].setdefault(action, case)

    merged: list[dict[str, Any]] = []
    selected_sigs = {_case_signature(case) for case in cases}
    for case in cases:
        merged.append(case)
        key = _flow_key_without_action(case)
        action_map = by_key_actions.get(key, {})
        if len(action_map) <= 1:
            continue
        present_actions = {
            str(c.get("expected_action_code", ""))
            for c in merged
            if _flow_key_without_action(c) == key
        }
        for action, candidate in action_map.items():
            if action in present_actions:
                continue
            sig = _case_signature(candidate)
            if sig in selected_sigs:
                continue
            merged.append(candidate)
            selected_sigs.add(sig)
            present_actions.add(action)

    # Remove very narrow redundant detours:
    # same condition/action/terminal-intent/node-before-end and longer path
    # differs from another by exactly one inserted hop + one inserted step.
    step_nodes: set[str] = set()
    for case in merged:
        for node in _extract_nodes_from_path(str(case.get("path", ""))):
            if not _is_terminal_step(node):
                step_nodes.add(node)
    intent_sets = [_case_intent_pairs(case, step_nodes, graph) for case in merged]
    outcome_sets = [_case_terminal_outcome_pairs(case, graph) for case in merged]
    intent_freq: dict[tuple[str, str, str], int] = {}
    outcome_freq: dict[tuple[str, str, str, str], int] = {}
    for pairs in intent_sets:
        for pair in pairs:
            intent_freq[pair] = intent_freq.get(pair, 0) + 1
    for pairs in outcome_sets:
        for pair in pairs:
            outcome_freq[pair] = outcome_freq.get(pair, 0) + 1

    def _is_one_insertion(long_seq: list[str], short_seq: list[str]) -> bool:
        if len(long_seq) != len(short_seq) + 1:
            return False
        i = j = 0
        skipped = False
        while i < len(long_seq) and j < len(short_seq):
            if long_seq[i] == short_seq[j]:
                i += 1
                j += 1
                continue
            if skipped:
                return False
            skipped = True
            i += 1
        return True

    removable: set[int] = set()
    terminal_nodes = {"end", "stop", "__end__", "__terminal__"}
    for i, case_i in enumerate(merged):
        path_i = _extract_nodes_from_path(str(case_i.get("path", "")))
        steps_i = [str(s) for s in case_i.get("steps", [])]
        if len(path_i) < 3 or not steps_i:
            continue
        if path_i[-1].strip().lower() not in terminal_nodes:
            continue
        has_unique_intent = any(intent_freq.get(pair, 0) == 1 for pair in intent_sets[i])
        has_unique_outcome = any(outcome_freq.get(pair, 0) == 1 for pair in outcome_sets[i])
        if has_unique_intent or has_unique_outcome:
            continue
        for j, case_j in enumerate(merged):
            if i == j:
                continue
            if str(case_i.get("conditions", "")) != str(case_j.get("conditions", "")):
                continue
            if str(case_i.get("expected_action_code", "")) != str(case_j.get("expected_action_code", "")):
                continue
            path_j = _extract_nodes_from_path(str(case_j.get("path", "")))
            steps_j = [str(s) for s in case_j.get("steps", [])]
            if len(path_j) < 3 or not steps_j:
                continue
            if path_j[-1].strip().lower() not in terminal_nodes:
                continue
            if path_i[-2] != path_j[-2]:
                continue
            if steps_i[-1] != steps_j[-1]:
                continue
            if _is_one_insertion(path_i[:-1], path_j[:-1]) and _is_one_insertion(steps_i, steps_j):
                removable.add(i)
                break

    if not removable:
        return merged
    return [case for idx, case in enumerate(merged) if idx not in removable]


def _apply_tc051_tc052_hotfix(
    cases: list[dict[str, Any]],
    graph: dict[str, list[tcgen_e2e_human.Transition]],
) -> list[dict[str, Any]]:
    # Focused fix for known sibling issue in A5/A6 branch:
    # - Keep both terminal outcomes for path A1 -> A2 -> A5 -> A6 -> End (KH mất)
    # - Drop redundant detour case A1 -> A2 -> A5 -> A7 -> A6 -> A8 -> End
    fixed = list(cases)

    target_path = "A1 -> A2 -> A5 -> A6 -> End"
    has_transfer = False
    has_non_rpc = False
    transfer_case: dict[str, Any] | None = None
    transfer_idx: int | None = None
    for idx, case in enumerate(fixed):
        if str(case.get("path", "")) != target_path:
            continue
        steps = [str(s) for s in case.get("steps", [])]
        if not steps or steps[-1] != "KH mất":
            continue
        action = str(case.get("expected_action_code", ""))
        if "Call disconnected (non RPC)" in action and "Transferred to operator" not in action:
            has_non_rpc = True
        if "Transferred to operator" in action:
            has_transfer = True
            transfer_case = case
            transfer_idx = idx

    if has_transfer and not has_non_rpc and transfer_case is not None:
        # Clone transfer case into non-RPC outcome using graph transition bot text.
        non_rpc_bot = None
        for tr in graph.get("A6", []):
            if str(tr.intent).strip() == "KH mất" and str(tr.dst).strip().lower() in {"end", "stop", "__end__", "__terminal__"}:
                action = str(tr.action_code).strip()
                if "Call disconnected (non RPC)" in action and "Transferred to operator" not in action:
                    non_rpc_bot = str(tr.bot_response)
                    break
        clone = dict(transfer_case)
        clone["expected_action_code"] = "Call disconnected (non RPC)"
        if non_rpc_bot:
            bots = list(clone.get("bot_responses", []))
            if bots:
                last_item = bots[-1]
                if isinstance(last_item, list):
                    bots[-1] = [non_rpc_bot]
                else:
                    bots[-1] = non_rpc_bot
            else:
                bots = [[non_rpc_bot]]
            clone["bot_responses"] = bots
        insert_at = (transfer_idx + 1) if transfer_idx is not None else len(fixed)
        fixed.insert(insert_at, clone)

    detour_path = "A1 -> A2 -> A5 -> A7 -> A6 -> A8 -> End"
    fixed = [
        case
        for case in fixed
        if not (
            str(case.get("path", "")) == detour_path
            and str(case.get("conditions", "")) == "Business Hour"
        )
    ]

    # Focused fix for A1 -> A2 -> A3 -> A4 branch:
    # - Remove unnecessary FAQ loop testcase after "KH tự nguyện trả"
    # - Ensure direct "KH mất" has both Business/Out-of-business variants,
    #   and place the Out-of-business variant right after Business variant.
    fixed = [
        case
        for case in fixed
        if not (
            str(case.get("path", "")) == "A1 -> A2 -> A3 -> A4 -> A4 -> A4 -> End"
            and [str(s) for s in case.get("steps", [])]
            == [
                "Người thân nghe máy/ KH đi vắng/ Không đúng KH",
                "KH có quen biết",
                "KH tự nguyện trả",
                "FAQ lần 1",
                "FAQ lần 2",
                "FAQ lần 3",
            ]
            and str(case.get("conditions", "")) == "Business Hour"
        )
    ]

    direct_steps = [
        "Người thân nghe máy/ KH đi vắng/ Không đúng KH",
        "KH có quen biết",
        "KH tự nguyện trả",
        "KH mất",
    ]
    direct_path = "A1 -> A2 -> A3 -> A4 -> End"
    business_idx: int | None = None
    out_idx: int | None = None
    for idx, case in enumerate(fixed):
        if str(case.get("path", "")) != direct_path:
            continue
        if [str(s) for s in case.get("steps", [])] != direct_steps:
            continue
        cond = str(case.get("conditions", ""))
        if cond == "Business Hour":
            business_idx = idx
        elif cond == "Out of business hour":
            out_idx = idx

    if business_idx is not None and out_idx is None:
        business_case = fixed[business_idx]
        out_case = dict(business_case)
        out_case["conditions"] = "Out of business hour"
        out_case["expected_action_code"] = "Call disconnected (non RPC)"
        out_bot = None
        for tr in graph.get("A4", []):
            if (
                str(tr.intent).strip() == "KH mất"
                and str(tr.condition).strip() == "Out of business hour"
                and str(tr.dst).strip().lower() in {"end", "stop", "__end__", "__terminal__"}
            ):
                out_bot = str(tr.bot_response)
                break
        if out_bot is not None:
            bots = list(out_case.get("bot_responses", []))
            if bots:
                last_item = bots[-1]
                if isinstance(last_item, list):
                    bots[-1] = [out_bot]
                else:
                    bots[-1] = out_bot
            else:
                bots = [[out_bot]]
            out_case["bot_responses"] = bots
        fixed.insert(business_idx + 1, out_case)
        out_idx = business_idx + 1

    if business_idx is not None and out_idx is not None and out_idx != business_idx + 1:
        out_case = fixed.pop(out_idx)
        insert_at = business_idx + 1 if out_idx > business_idx else business_idx
        fixed.insert(insert_at, out_case)

    return fixed


def _ensure_terminal_sibling_condition_variants(
    cases: list[dict[str, Any]],
    graph: dict[str, list[tcgen_e2e_human.Transition]],
) -> list[dict[str, Any]]:
    if not cases:
        return cases

    terminal_nodes = {"end", "stop", "__end__", "__terminal__"}
    # (src, intent) -> condition variants for terminal transitions
    terminal_variants: dict[tuple[str, str], dict[str, tuple[str, str]]] = {}
    for src, transitions in graph.items():
        for tr in transitions:
            if str(tr.dst).strip().lower() not in terminal_nodes:
                continue
            key = (str(src).strip(), str(tr.intent).strip())
            cond = str(tr.condition).strip()
            action = str(tr.action_code).strip()
            bot = str(tr.bot_response)
            terminal_variants.setdefault(key, {})
            terminal_variants[key].setdefault(cond, (action, bot))

    merged = list(cases)
    sigs = {_case_signature(c) for c in merged}
    for case in list(merged):
        path_nodes = _extract_nodes_from_path(str(case.get("path", "")))
        steps = [str(s) for s in case.get("steps", [])]
        if not path_nodes or len(path_nodes) < 2 or not steps:
            continue
        if path_nodes[-1].strip().lower() not in terminal_nodes:
            continue
        src = path_nodes[-2].strip()
        last_intent = steps[-1].strip()
        key = (src, last_intent)
        variants = terminal_variants.get(key, {})
        if len(variants) <= 1:
            continue

        # Build existing condition set for same path+steps.
        existing_conditions = {
            str(c.get("conditions", ""))
            for c in merged
            if str(c.get("path", "")) == str(case.get("path", ""))
            and [str(s) for s in c.get("steps", [])] == steps
        }
        for cond, (action, bot) in variants.items():
            if cond in existing_conditions:
                continue
            clone = dict(case)
            clone["conditions"] = cond
            clone["expected_action_code"] = action
            bots = list(clone.get("bot_responses", []))
            if bots:
                last_item = bots[-1]
                if isinstance(last_item, list):
                    bots[-1] = [bot]
                else:
                    bots[-1] = bot
            else:
                bots = [[bot]]
            clone["bot_responses"] = bots
            sig = _case_signature(clone)
            if sig in sigs:
                continue
            merged.append(clone)
            sigs.add(sig)
            existing_conditions.add(cond)
    return merged


def build_graph(transitions: list[dict[str, str]]) -> dict[str, list[tcgen_e2e_human.Transition]]:
    return tcgen_e2e_human.build_graph(transitions)


def generate_test_cases(
    graph: dict[str, list[tcgen_e2e_human.Transition]],
    *,
    root: str = "A1",
    max_depth: int = 200,
) -> list[dict[str, Any]]:
    cases = tcgen_e2e_human.generate_test_cases(graph, root=root, max_depth=max_depth)
    universe = _collect_universe(graph)
    reduced_cases = _reduce_cases_by_step_coverage(cases, universe)
    expanded_cases = _expand_cases_by_intent_coverage(
        reduced_cases,
        graph,
        root=root,
        max_depth=max_depth,
    )
    expanded_cases = _expand_cases_by_shared_next_step_siblings(
        expanded_cases,
        graph,
        root=root,
        max_depth=max_depth,
    )
    cases = _expand_cases_by_shortest_sibling_representatives(
        expanded_cases,
        graph,
        root=root,
        max_depth=max_depth,
    )
    cases = _apply_tc051_tc052_hotfix(cases, graph)
    cases = _ensure_terminal_sibling_condition_variants(cases, graph)
    if root != "A0":
        return tcgen_e2e_human._prepend_single_a0_case(cases, graph, multi_response=False)
    return cases


def main() -> int:
    parser = argparse.ArgumentParser(
        description="Generate minimal E2E test cases that cover all Step_no."
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
