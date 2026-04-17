"""
Generate realistic customer conversation utterances for each test case using a hybrid approach:
1. First try TLS Client Bot (NLP intent classifier)
2. If TLS bot fails for a specific step, fallback to Llama 3.1 (Ollama) for that step only

Multi‑intent steps are split by " \ " (backslash) and multiple rounds are generated:
- Round N selects the N-th intent from each step (circularly if a step has fewer intents).
- Bot responses may contain multiple variants separated by " \ "; one is chosen randomly per generation.
"""

from __future__ import annotations

from collections.abc import Iterable
import argparse
import json
import os
import random
import re
import sys
from contextlib import contextmanager
from pathlib import Path
from typing import Any
import unicodedata

import openpyxl
from openpyxl.styles import Alignment
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.worksheet import Worksheet

# Try to import Ollama (optional dependency)
try:
    from ollama import chat as ollama_chat
except ImportError:
    ollama_chat = None

# TLS Client Bot imports
DEFAULT_MODEL = "tls_client_bot"
MAX_WORDS_PER_RESPONSE = 12

PROJECT_ROOT = Path(__file__).resolve().parent
TLS_BOT_DIR = PROJECT_ROOT / "tls_client_bot"
TLS_INTENTS_PATH = TLS_BOT_DIR / "intents.json"

_TLS_GET_RESPONSE = None
_TLS_INTENTS_DATA: dict[str, Any] | None = None
_TAGS_SET: set[str] | None = None
_PATTERN_TO_TAG: dict[str, str] | None = None
_PATTERN_TO_TAG_FOLDED: dict[str, str] | None = None

# Llama 3.1 model name for fallback
LLAMA_MODEL = "llama3.1"


@contextmanager
def _temporary_chdir(path: Path):
    old = Path.cwd()
    os.chdir(path)
    try:
        yield
    finally:
        os.chdir(old)


def _load_tls_intents() -> dict[str, Any]:
    global _TLS_INTENTS_DATA, _TAGS_SET, _PATTERN_TO_TAG, _PATTERN_TO_TAG_FOLDED
    if _TLS_INTENTS_DATA is not None:
        return _TLS_INTENTS_DATA

    if not TLS_INTENTS_PATH.exists():
        raise FileNotFoundError(f"Missing tls bot intents.json: {TLS_INTENTS_PATH}")

    data = json.loads(TLS_INTENTS_PATH.read_text(encoding="utf-8"))
    _TLS_INTENTS_DATA = data
    tags_set: set[str] = set()
    pattern_to_tag: dict[str, str] = {}
    pattern_to_tag_folded: dict[str, str] = {}

    for intent in data.get("intents", []):
        tag = str(intent.get("tag", "")).strip()
        if not tag:
            continue
        tags_set.add(tag)
        for pattern in intent.get("patterns", []) or []:
            p = str(pattern).strip().lower()
            if p and p not in pattern_to_tag:
                pattern_to_tag[p] = tag
            folded_p = _fold_for_match(p)
            if folded_p and folded_p not in pattern_to_tag_folded:
                pattern_to_tag_folded[folded_p] = tag

    _TAGS_SET = tags_set
    _PATTERN_TO_TAG = pattern_to_tag
    _PATTERN_TO_TAG_FOLDED = pattern_to_tag_folded
    return data


def _load_tls_get_response():
    """
    Lazily import `tls_client_bot/chat.py` and expose its `get_response()` function.

    Note: tls_client_bot code uses non-package imports like `from chat import ...`,
    and loads `intents.json`/`data.pth` via relative paths; we handle this by
    temporarily adjusting sys.path + working directory.
    """

    global _TLS_GET_RESPONSE
    if _TLS_GET_RESPONSE is not None:
        return _TLS_GET_RESPONSE

    if not TLS_BOT_DIR.exists():
        _TLS_GET_RESPONSE = None
        return None

    try:
        # Make `chat.py`, `model.py`, etc importable as top-level modules.
        tls_str = str(TLS_BOT_DIR)
        if tls_str not in sys.path:
            sys.path.insert(0, tls_str)

        with _temporary_chdir(TLS_BOT_DIR):
            # Import as a top-level module name expected by tls_client_bot.
            from chat import get_response as tls_get_response  # type: ignore

        _TLS_GET_RESPONSE = tls_get_response
        return _TLS_GET_RESPONSE
    except Exception:
        # Most commonly: missing torch/nltk deps, or NLTK data download failing.
        _TLS_GET_RESPONSE = None
        return None


def _extract_latest_attempt(raw_step: str) -> tuple[str | None, str | None]:
    """
    Trích xuất attempt number cao nhất/cuối cùng từ step.
    Returns: (full_attempt_text, attempt_number)
    Ví dụ: "lần 1" -> ("lần 1", "1"), "lần 2" -> ("lần 2", "2")
    """
    attempt_matches = re.finditer(r"\b(lần\s*(\d+))\b", raw_step, flags=re.IGNORECASE)
    attempts = []
    for match in attempt_matches:
        attempts.append({
            'full': match.group(1),
            'number': int(match.group(2)),
            'position': match.start()
        })
    
    if attempts:
        # Sắp xếp theo vị trí và lấy attempt cuối cùng (thường là số cao nhất)
        attempts.sort(key=lambda x: x['position'])
        latest = attempts[-1]
        return latest['full'], str(latest['number'])
    return None, None


def _normalize_step_to_bot_message(step: str) -> str:
    """
    Turn a tcgen step string into a message that `tls_client_bot` can classify.
    Steps now use backslash as delimiter. We take the first intent and preserve
    the highest attempt number.
    """
    _load_tls_intents()
    tags_set = _TAGS_SET or set()

    raw = (step or "").strip()
    if not raw:
        return raw

    # Split on backslash (with optional spaces)
    parts = [p.strip() for p in re.split(r"\s*\\\s*", raw) if p.strip()]
    first_part = parts[0] if parts else raw
    is_multi_intent = len(parts) > 1
    first_folded = _fold_for_match(first_part)

    # 1) Special cases: im lặng, không nghe rõ
    if "im lang" in first_folded:
        for tag in tags_set:
            if "im lang" in _fold_for_match(tag):
                return tag

    if "khong nghe ro" in first_folded:
        for tag in tags_set:
            if "khong nghe ro" in _fold_for_match(tag):
                return tag

    # 2) Multi-intent: take first intent, keep attempt number from the whole step
    if is_multi_intent:
        selected_intent = parts[0]
        attempt_full, _ = _extract_latest_attempt(raw)
        if attempt_full:
            # Remove any existing attempt from selected_intent
            selected_intent = re.sub(r"\s*lần\s*\d+\s*", "", selected_intent, flags=re.IGNORECASE).strip()
            return f"{selected_intent} {attempt_full}".strip()
        return first_part

    # 3) Single intent: try exact tag match
    for tag in tags_set:
        tag_folded = _fold_for_match(tag)
        if tag_folded == first_folded:
            return tag
        # If tag is like "A / B", first part of tag might match
        tag_first = tag_folded.split("/")[0].strip()
        if tag_first == first_folded:
            return tag

    # 4) Fallback
    return first_part


_BANNED_CHARS_RE = re.compile(r"[\*\^\$\#\@\%\&\?\!]")
_TRAILING_PUNCT_RE = re.compile(r"[\.!\?]+$")
_MULTISPACE_RE = re.compile(r"\s+")


def _sanitize_utterance(text: str) -> str:
    """
    Apply the same output constraints style as `generate_test_data.py`:
    - all lowercase
    - remove some special characters
    - no trailing .?! characters
    - no commas
    - max word count per line
    """
    if not text:
        return "(không sinh được)"

    t = str(text).strip()

    # If model couldn't classify, keep a consistent placeholder.
    if "do not understand" in t.lower():
        return "(không sinh được)"

    t = t.replace(",", " ")
    t = _BANNED_CHARS_RE.sub("", t)
    t = t.replace("\n", " ")
    t = _TRAILING_PUNCT_RE.sub("", t).strip()
    t = t.lower()
    t = _MULTISPACE_RE.sub(" ", t).strip()

    words = t.split(" ")
    if len(words) > MAX_WORDS_PER_RESPONSE:
        t = " ".join(words[:MAX_WORDS_PER_RESPONSE]).strip()

    return t if t else "(không sinh được)"


def _fold_for_match(s: str) -> str:
    """
    Case-fold + accent-fold (Vietnamese) for robust comparisons.
    """
    base = unicodedata.normalize("NFD", s.lower())
    return "".join(ch for ch in base if unicodedata.category(ch) != "Mn")


def _build_llama_prompt_for_step(step: str, step_number: int, total_steps: int, context: list[str], expected_bot_response: str = "") -> str:
    """
    Build prompt for Llama 3.1 to generate a single utterance for a specific step.
    Includes context from previous steps for better continuity.
    """
    context_text = ""
    if context:
        context_lines = []
        for i, ctx in enumerate(context, 1):
            context_lines.append(f"{i}. {ctx}")
        context_text = "Các câu trước đó:\n" + "\n".join(context_lines) + "\n\n"
    
    bot_response_text = ""
    if expected_bot_response:
        bot_response_text = f"\nBot sẽ đáp: \"{expected_bot_response}\""
    
    return f"""Bạn là KHÁCH HÀNG (KH) đang nói chuyện TRỰC TIẾP qua điện thoại với tổng đài viên về KHOẢN VAY.

NHIỆM VỤ: chỉ tạo 1 câu lời thoại KH cho bước số {step_number} dựa trên kịch bản.

{context_text}Kịch bản bước {step_number}: {step}{bot_response_text}

YÊU CẦU BẮT BUỘC:
- Chỉ tạo DUY NHẤT 1 câu thoại KH đúng với mõi intent tương ứng
- CHỈ là lời KH nói, không phải mô tả tình huống
- KHÔNG được nhắc lại hoặc diễn giải kịch bản
- KHÔNG được viết kiểu: "khách hàng im lặng", "người thân nói", "cuộc gọi..."
- Nếu tình huống là "im lặng" → viết câu tự nhiên thể hiện không nghe rõ (vd: alo em không nghe rõ)
- Nếu "người thân nghe hộ" → vẫn phải nói như người đang nghe điện thoại (vd: chị là mẹ chau đang nghe)

FORMAT CỨNG:
- Chỉ output 1 câu duy nhất, không đánh số
- KHÔNG thêm tiêu đề, KHÔNG markdown, KHÔNG giải thích
- KHÔNG ký tự đặc biệt: *^$#@%&?!
- KHÔNG dấu chấm, chấm hỏi, chấm than cuối câu
- KHÔNG dấu phẩy
- viết thường toàn bộ
- tối đa {MAX_WORDS_PER_RESPONSE} từ

VÍ DỤ:
Kịch bản: KH im lặng lần 1
Bot sẽ đáp: "alo em nghe không rõ ạ"
Output: alo em nghe không rõ anh nói lại giúp em với

Đầu ra:"""


def _generate_single_with_llama(step: str, step_number: int, total_steps: int, context: list[str], expected_bot_response: str = "") -> str:
    """Generate a single utterance using Llama 3.1 via Ollama."""
    if not step:
        return "(không sinh được)"
    
    if ollama_chat is None:
        return "(không sinh được)"
    
    prompt = _build_llama_prompt_for_step(step, step_number, total_steps, context, expected_bot_response)
    try:
        response = ollama_chat(
            model=LLAMA_MODEL,
            messages=[{"role": "user", "content": prompt}],
        )
        content = (response.get("message") or {}).get("content") or ""
        # Clean up the response
        content = content.strip()
        # Remove any numbering if present
        content = re.sub(r"^\s*\d+[.)]\s*", "", content)
        content = _sanitize_utterance(content)
        return content if content else "(không sinh được)"
    except Exception as e:
        print(f"    [WARN] Llama generation failed for step {step_number}: {e}")
        return "(không sinh được)"


def _fallback_response_for_message(msg: str) -> str:
    """Fallback when TLS bot fails: use intents.json pattern matching."""
    _load_tls_intents()
    intents = _TLS_INTENTS_DATA.get("intents", []) if _TLS_INTENTS_DATA else []
    tags_set = _TAGS_SET or set()
    pattern_to_tag = _PATTERN_TO_TAG or {}
    pattern_to_tag_folded = _PATTERN_TO_TAG_FOLDED or {}

    m = (msg or "").strip()
    if not m:
        return "(không sinh được)"

    m_lower = m.lower()
    m_folded = _fold_for_match(m_lower)

    # Remove attempt number for matching
    m_without_attempt = re.sub(r'\s*lần\s*\d+\s*', ' ', m_lower).strip()
    m_without_attempt_folded = _fold_for_match(m_without_attempt)

    tag: str | None = None
    for t in tags_set:
        t_folded = _fold_for_match(t)
        t_without_attempt = re.sub(r'\s*lần\s*\d+\s*', '', t_folded).strip()
        if t_folded == m_folded or t_without_attempt == m_without_attempt_folded:
            tag = t
            break
        t_first = t_folded.split("/")[0].strip()
        if t_first == m_folded or t_first == m_without_attempt_folded:
            tag = t
            break

    if not tag:
        tag = pattern_to_tag.get(m_lower) or pattern_to_tag.get(m_without_attempt)
    if not tag:
        tag = pattern_to_tag_folded.get(m_folded) or pattern_to_tag_folded.get(m_without_attempt_folded)

    if not tag:
        return "(không sinh được)"

    for intent in intents:
        if str(intent.get("tag", "")).strip() == tag:
            responses = intent.get("responses") or []
            if responses:
                selected = random.choice([str(r) for r in responses if str(r).strip()])
                return selected
    return "(không sinh được)"


def _generate_step_with_tls_bot(step: str) -> str:
    """Generate a single utterance using TLS Client Bot."""
    tls_get_response = _load_tls_get_response()
    
    msg = _normalize_step_to_bot_message(step)
    
    if tls_get_response is None:
        resp = _fallback_response_for_message(msg)
        return _sanitize_utterance(resp)
    
    try:
        resp = tls_get_response(msg, context=None, token=None)
        if not resp or "do not understand" in str(resp).lower():
            resp = _fallback_response_for_message(msg)
    except Exception:
        resp = _fallback_response_for_message(msg)
    
    return _sanitize_utterance(resp)


def _pick_random_bot_response(bot_response_str: str) -> str:
    """
    If bot_response_str contains multiple responses separated by " \ ",
    pick one at random. Otherwise return the string as is.
    """
    if not bot_response_str:
        return ""
    # Split on backslash with optional spaces
    parts = [p.strip() for p in re.split(r"\s*\\\s*", bot_response_str) if p.strip()]
    if not parts:
        return bot_response_str
    return random.choice(parts)


def _generate_hybrid_for_round(steps: list[str], bot_responses: list[str], tc_id: str, round_num: int) -> str:
    """
    Hybrid generation for a single round (specific selected intents per step).
    Steps already contain the specific intent (e.g., after circular selection).
    For each step, if the bot_response string contains multiple variants (separated by " \ "),
    we randomly pick one to use as expected bot response.
    """
    print(f"\n[DEBUG] Processing TC: {tc_id}, Round {round_num} with hybrid approach")
    
    utterances: list[str] = []
    tls_failed_steps: list[int] = []
    
    # First pass: try TLS bot for all steps
    for i, step in enumerate(steps, 1):
        tls_result = _generate_step_with_tls_bot(step)
        utterances.append(tls_result)
        
        if tls_result == "(không sinh được)":
            tls_failed_steps.append(i - 1)  # Store 0-based index
            print(f"  Step {i}: TLS bot failed -> will use Llama fallback")
        else:
            print(f"  Step {i}: TLS bot succeeded -> '{tls_result}'")
    
    # Second pass: for failed steps, use Llama with context from previous utterances
    for idx in tls_failed_steps:
        step_number = idx + 1
        step = steps[idx]
        # Pick a random bot response variant if multiple exist
        raw_bot_resp = bot_responses[idx] if idx < len(bot_responses) else ""
        expected_bot_response = _pick_random_bot_response(raw_bot_resp)
        
        # Get context from previous utterances (excluding failed ones that are not yet generated)
        context = []
        for j in range(idx):
            if j < len(utterances) and utterances[j] != "(không sinh được)":
                context.append(utterances[j])
        
        print(f"  Step {step_number}: Generating with Llama (context: {len(context)} previous utterances)")
        
        llama_result = _generate_single_with_llama(step, step_number, len(steps), context, expected_bot_response)
        utterances[idx] = llama_result
        print(f"    Llama result: '{llama_result}'")
    
    # Format output
    result = "\n".join(f"{i+1}. {p}" for i, p in enumerate(utterances[:len(steps)]))
    
    tls_success = len(steps) - len(tls_failed_steps)
    llama_used = len(tls_failed_steps)
    print(f"  Summary Round {round_num}: {tls_success}/{len(steps)} steps from TLS, {llama_used} steps from Llama")
    
    return result


def _parse_step_intents(step_str: str) -> list[str]:
    """Split a step string into a list of intents using backslash delimiter."""
    return [p.strip() for p in re.split(r"\s*\\\s*", step_str) if p.strip()]


def _auto_fit_column(ws: Worksheet, col_idx: int, extra: int = 2) -> None:
    col_letter = get_column_letter(col_idx)
    max_len = 0
    for cell in ws[col_letter]:
        if cell.value:
            for part in str(cell.value).split("\n"):
                max_len = max(max_len, len(part))
    ws.column_dimensions[col_letter].width = max_len + extra


def fill_excel_from_json(
    json_path: Path,
    excel_path: Path,
    model: str = DEFAULT_MODEL,
    limit: int | None = None,
) -> None:
    """
    Load TC JSON, generate test data for each test case, producing one row per round.
    Handles multi‑intent steps split by backslash.
    """
    data = json.loads(json_path.read_text(encoding="utf-8"))
    if not isinstance(data, list):
        raise ValueError("Input JSON must be a list of test cases.")

    cases: list[dict[str, Any]] = []
    for i, item in enumerate(data):
        if not isinstance(item, dict):
            continue
        tc_id = str(item.get("tc_id", f"TC_{i+1:03d}"))
        steps = item.get("steps", [])
        if not isinstance(steps, Iterable):
            steps = []
        bot_responses = item.get("bot_responses", [])
        if not isinstance(bot_responses, Iterable):
            bot_responses = []

        clean_steps = [
            str(s).strip()
            for s in steps
            if str(s).strip() and not str(s).strip().startswith("(cycle to")
        ]
        clean_bot_responses = [
            str(r).strip()
            for r in bot_responses
            if str(r).strip()
        ]
        
        # Ensure bot_responses length matches steps length
        while len(clean_bot_responses) < len(clean_steps):
            clean_bot_responses.append("")
        
        expected_action_code = str(item.get("expected_action_code", ""))
        tc_path = str(item.get("path", ""))

        # Parse each step into list of intents (using backslash)
        steps_intents = [_parse_step_intents(s) for s in clean_steps]
        if not steps_intents:
            continue
        max_rounds = max(len(intents) for intents in steps_intents) if steps_intents else 1

        cases.append({
            "tc_id": tc_id,
            "steps": clean_steps,
            "steps_intents": steps_intents,
            "bot_responses": clean_bot_responses,
            "expected_action_code": expected_action_code,
            "path": tc_path,
            "max_rounds": max_rounds,
        })

    if limit is not None:
        cases = cases[:limit]

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "TestCases"
    ws["A1"] = "TC_ID"
    ws["B1"] = "Round"
    ws["C1"] = "Test Scenario"
    ws["D1"] = "Bot Responses"
    ws["E1"] = "Path"
    ws["F1"] = "Expected Action Code"
    ws["G1"] = "Test Data"

    row_idx = 2
    for case in cases:
        tc_id = case["tc_id"]
        steps_intents = case["steps_intents"]
        bot_responses = case["bot_responses"]
        expected_action_code = case["expected_action_code"]
        tc_path = case["path"]
        max_rounds = case["max_rounds"]

        for round_num in range(1, max_rounds + 1):
            round_steps = []
            for i, intents in enumerate(steps_intents):
                idx = (round_num - 1) % len(intents)
                round_steps.append(intents[idx])
            
            scenario_text = "\n".join(f"{i+1}. {step}" for i, step in enumerate(round_steps))
            bot_responses_text = "\n".join(f"{i+1}. {r}" for i, r in enumerate(bot_responses) if r)
            
            test_data = _generate_hybrid_for_round(round_steps, bot_responses, tc_id, round_num)

            ws.cell(row=row_idx, column=1, value=tc_id)
            ws.cell(row=row_idx, column=2, value=round_num)
            cell_c = ws.cell(row=row_idx, column=3, value=scenario_text)
            cell_c.alignment = Alignment(wrap_text=True)
            cell_d = ws.cell(row=row_idx, column=4, value=bot_responses_text)
            cell_d.alignment = Alignment(wrap_text=True)
            ws.cell(row=row_idx, column=5, value=tc_path)
            ws.cell(row=row_idx, column=6, value=expected_action_code)
            cell_g = ws.cell(row=row_idx, column=7, value=test_data)
            cell_g.alignment = Alignment(wrap_text=True)
            print(f"  Generated: {tc_id} Round {round_num} ({len(round_steps)} steps)")

            row_idx += 1

    for c in range(1, 8):
        _auto_fit_column(ws, c, extra=4)

    excel_path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(excel_path)


def fill_excel_from_excel(
    excel_path: Path,
    out_path: Path,
    model: str = DEFAULT_MODEL,
    limit: int | None = None,
) -> None:
    """
    Read existing Excel, fill Test Data column using hybrid approach, one row per round.
    Assumes the input Excel already has columns: TC_ID, Test Scenario, Bot Responses, Path, Expected Action Code.
    """
    wb = openpyxl.load_workbook(excel_path)
    ws = wb.active
    max_row = ws.max_row
    if limit is not None:
        max_row = min(max_row, 1 + limit)

    # Find columns by header names
    header_to_col: dict[str, int] = {}
    for col in range(1, ws.max_column + 1):
        header_val = ws.cell(row=1, column=col).value
        if header_val is None:
            continue
        header_to_col[str(header_val).strip()] = col

    tc_id_col = header_to_col.get("TC_ID", 1)
    scenario_col = header_to_col.get("Test Scenario", 2)
    bot_responses_col = header_to_col.get("Bot Responses", 3)
    path_col = header_to_col.get("Path", 4)
    expected_action_col = header_to_col.get("Expected Action Code", 5)

    rows_data = []
    for row in range(2, max_row + 1):
        tc_id = ws.cell(row=row, column=tc_id_col).value or ""
        scenario_cell = ws.cell(row=row, column=scenario_col).value or ""
        bot_responses_cell = ws.cell(row=row, column=bot_responses_col).value or "" if bot_responses_col else ""
        path_val = ws.cell(row=row, column=path_col).value or ""
        expected_action = ws.cell(row=row, column=expected_action_col).value or ""

        # Parse steps from scenario (numbered lines)
        steps: list[str] = []
        numbered_re = re.compile(r"^\s*(\d+)[.)]\s*(.*)$")
        for raw_line in str(scenario_cell).strip().split("\n"):
            line = raw_line.strip()
            if not line:
                continue
            m = numbered_re.match(raw_line)
            if m:
                cleaned = (m.group(2) or "").strip()
                if cleaned:
                    steps.append(cleaned)
                else:
                    steps.append("")
            elif steps:
                steps[-1] = f"{steps[-1].rstrip()} {line}".strip()
            else:
                steps.append(line)

        # Parse bot responses similarly
        bot_responses: list[str] = []
        for raw_line in str(bot_responses_cell).strip().split("\n"):
            line = raw_line.strip()
            if not line:
                continue
            m = numbered_re.match(raw_line)
            if m:
                cleaned = (m.group(2) or "").strip()
                if cleaned:
                    bot_responses.append(cleaned)
                else:
                    bot_responses.append("")
            elif bot_responses:
                bot_responses[-1] = f"{bot_responses[-1].rstrip()} {line}".strip()

        while len(bot_responses) < len(steps):
            bot_responses.append("")
        if not steps:
            continue

        rows_data.append({
            "tc_id": tc_id,
            "steps": steps,
            "bot_responses": bot_responses,
            "path": path_val,
            "expected_action": expected_action,
        })

    # Build new workbook with Round column
    new_wb = openpyxl.Workbook()
    new_ws = new_wb.active
    new_ws.title = "TestCases"
    new_ws["A1"] = "TC_ID"
    new_ws["B1"] = "Round"
    new_ws["C1"] = "Test Scenario"
    new_ws["D1"] = "Bot Responses"
    new_ws["E1"] = "Path"
    new_ws["F1"] = "Expected Action Code"
    new_ws["G1"] = "Test Data"

    row_idx = 2
    for case in rows_data:
        tc_id = case["tc_id"]
        steps = case["steps"]
        bot_responses = case["bot_responses"]
        # Parse each step into intents using backslash
        steps_intents = [_parse_step_intents(s) for s in steps]
        if not steps_intents:
            continue
        max_rounds = max(len(intents) for intents in steps_intents)
        for round_num in range(1, max_rounds + 1):
            round_steps = []
            for i, intents in enumerate(steps_intents):
                idx = (round_num - 1) % len(intents)
                round_steps.append(intents[idx])
            scenario_text = "\n".join(f"{i+1}. {step}" for i, step in enumerate(round_steps))
            bot_responses_text = "\n".join(f"{i+1}. {r}" for i, r in enumerate(bot_responses) if r)
            test_data = _generate_hybrid_for_round(round_steps, bot_responses, str(tc_id), round_num)

            new_ws.cell(row=row_idx, column=1, value=tc_id)
            new_ws.cell(row=row_idx, column=2, value=round_num)
            cell_c = new_ws.cell(row=row_idx, column=3, value=scenario_text)
            cell_c.alignment = Alignment(wrap_text=True)
            cell_d = new_ws.cell(row=row_idx, column=4, value=bot_responses_text)
            cell_d.alignment = Alignment(wrap_text=True)
            new_ws.cell(row=row_idx, column=5, value=case["path"])
            new_ws.cell(row=row_idx, column=6, value=case["expected_action"])
            cell_g = new_ws.cell(row=row_idx, column=7, value=test_data)
            cell_g.alignment = Alignment(wrap_text=True)
            print(f"  Generated: {tc_id} Round {round_num} ({len(round_steps)} steps)")
            row_idx += 1

    for c in range(1, 8):
        _auto_fit_column(new_ws, c, extra=4)

    out_path.parent.mkdir(parents=True, exist_ok=True)
    new_wb.save(out_path)


def main() -> int:
    parser = argparse.ArgumentParser(
        description="Generate test data using hybrid approach: TLS Client Bot + Llama 3.1 fallback per step, with multi‑intent step handling (backslash delimiter) and multiple rounds."
    )
    parser.add_argument(
        "--in",
        dest="in_path",
        type=Path,
        required=True,
        help="Input: JSON (tcgen output) or Excel (tc_to_excel output).",
    )
    parser.add_argument(
        "--out",
        dest="out_path",
        type=Path,
        default=None,
        help="Output Excel path. If input is Excel, defaults to overwriting input.",
    )
    parser.add_argument(
        "--model",
        type=str,
        default=DEFAULT_MODEL,
        help=f"Kept for compatibility (default: {DEFAULT_MODEL}).",
    )
    parser.add_argument(
        "--limit",
        type=int,
        default=None,
        help="Process only the first N test cases (for testing).",
    )
    args = parser.parse_args()

    in_path = args.in_path
    if not in_path.exists():
        print(f"Error: Input file not found: {in_path}")
        return 1

    out_path = args.out_path
    if out_path is None:
        out_path = in_path if in_path.suffix.lower() == ".xlsx" else Path("output/test_cases.xlsx")

    print("=" * 60)
    print("HYBRID GENERATION MODE: TLS Client Bot + Llama 3.1 Fallback (per-step)")
    print("Multi‑intent steps: each round selects the N-th intent (circularly).")
    print("Bot responses may have multiple variants (separated by ' \\ ') -> random selection.")
    print("=" * 60)
    
    if ollama_chat is None:
        print("⚠️  Warning: Ollama not installed. Llama fallback will not work.")
        print("   Install with: pip install ollama")
        print("   Or run: ollama pull llama3.1")
        print()

    if in_path.suffix.lower() == ".json":
        print(f"Reading JSON: {in_path}")
        print(f"Writing Excel with test data: {out_path}")
        fill_excel_from_json(in_path, out_path, model=args.model, limit=args.limit)
    else:
        print(f"Reading Excel: {in_path}")
        print("Generating multiple rounds (one per intent combination)...")
        fill_excel_from_excel(in_path, out_path, model=args.model, limit=args.limit)

    print(f"\nDone. Output: {out_path}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())