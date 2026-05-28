from __future__ import annotations

import json
import shutil
from datetime import datetime
from pathlib import Path
from typing import Any, Literal

PROMPTS_DIR = Path(__file__).resolve().parent / "prompts"
LLM_FILE = PROMPTS_DIR / "llm_full_tc.txt"
HYBRID_FILE = PROMPTS_DIR / "hybrid_step.txt"
META_FILE = PROMPTS_DIR / "meta.json"

PromptKind = Literal["llm", "hybrid"]

DEFAULT_LLM_TEMPLATE = """Bạn là KHÁCH HÀNG (KH) đang nói chuyện TRỰC TIẾP qua điện thoại với tổng đài viên về KHOẢN VAY.

NHIỆM VỤ: chỉ tạo lời thoại KH (KHÔNG mô tả, KHÔNG giải thích, KHÔNG kể chuyện).

TC_ID: {tc_id}

Kịch bản (mỗi bước có thể chứa nhiều intent cách nhau bằng " \\ " – hãy lấy INTENT ĐẦU TIÊN để viết):
{steps_text}

YÊU CẦU BẮT BUỘC:
- Mỗi bước = 1 câu thoại KH tương ứng đúng thứ tự
- Nếu 1 dòng có nhiều intent (" \\ ") → CHỌN 1 INTENT ĐẦU TIÊN ĐỂ VIẾT
- Tổng số câu = {step_count}
- CHỈ là lời KH nói, không phải mô tả tình huống
- KHÔNG được nhắc lại hoặc diễn giải kịch bản
- KHÔNG được viết kiểu: "khách hàng im lặng", "người thân nói", "cuộc gọi..."
- Nếu tình huống là "im lặng" → viết câu tự nhiên thể hiện không nghe rõ (vd: alo em không nghe rõ)
- Nếu "người thân nghe hộ" → vẫn phải nói như người đang nghe điện thoại (vd: chị là mẹ chau đang nghe)

FORMAT CỨNG:
- Chỉ output danh sách đánh số (1., 2., 3...)
- Mỗi dòng 1 câu duy nhất
- KHÔNG thêm tiêu đề, KHÔNG markdown, KHÔNG giải thích
- KHÔNG ký tự đặc biệt: *^$#@%&?!
- KHÔNG dấu chấm, chấm hỏi, chấm than cuối câu
- KHÔNG dấu phẩy
- viết thường toàn bộ
- tối đa {max_words} từ mỗi câu

VÍ DỤ:

Input:
- KH im lặng lần 1
- KH im lặng lần 2

Output:
1. alo em nghe không rõ anh nói lại giúp em với
2. dạ vẫn chưa nghe rõ anh nói lớn hơn chút được không

Đầu ra:"""

DEFAULT_HYBRID_TEMPLATE = """Bạn là KHÁCH HÀNG (KH) đang nói chuyện TRỰC TIẾP qua điện thoại với tổng đài viên về KHOẢN VAY.

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
- tối đa {max_words} từ

VÍ DỤ:
Kịch bản: KH im lặng lần 1
Bot sẽ đáp: "alo em nghe không rõ ạ"
Output: alo em nghe không rõ anh nói lại giúp em với

Đầu ra:"""

PLACEHOLDERS: dict[PromptKind, list[str]] = {
    "llm": ["tc_id", "steps_text", "step_count", "max_words"],
    "hybrid": [
        "step_number",
        "total_steps",
        "context_text",
        "step",
        "bot_response_text",
        "max_words",
    ],
}


def _path_for(kind: PromptKind) -> Path:
    return LLM_FILE if kind == "llm" else HYBRID_FILE


def _default_for(kind: PromptKind) -> str:
    return DEFAULT_LLM_TEMPLATE if kind == "llm" else DEFAULT_HYBRID_TEMPLATE


def ensure_prompt_files() -> None:
    PROMPTS_DIR.mkdir(parents=True, exist_ok=True)
    if not LLM_FILE.exists():
        LLM_FILE.write_text(DEFAULT_LLM_TEMPLATE, encoding="utf-8")
    if not HYBRID_FILE.exists():
        HYBRID_FILE.write_text(DEFAULT_HYBRID_TEMPLATE, encoding="utf-8")
    if not META_FILE.exists():
        META_FILE.write_text(
            json.dumps(
                {
                    "llm": {"updated_at": None},
                    "hybrid": {"updated_at": None},
                },
                indent=2,
            ),
            encoding="utf-8",
        )


def get_prompt(kind: PromptKind) -> dict[str, Any]:
    ensure_prompt_files()
    path = _path_for(kind)
    meta = json.loads(META_FILE.read_text(encoding="utf-8"))
    content = path.read_text(encoding="utf-8")
    return {
        "kind": kind,
        "content": content,
        "placeholders": PLACEHOLDERS[kind],
        "updated_at": meta.get(kind, {}).get("updated_at"),
        "path": str(path.name),
    }


def save_prompt(kind: PromptKind, content: str) -> dict[str, Any]:
    ensure_prompt_files()
    path = _path_for(kind)
    if path.exists():
        stamp = datetime.utcnow().strftime("%Y%m%d_%H%M%S")
        backup = PROMPTS_DIR / f"{path.stem}_backup_{stamp}.txt"
        shutil.copy2(path, backup)
    path.write_text(content, encoding="utf-8")
    meta = json.loads(META_FILE.read_text(encoding="utf-8"))
    meta.setdefault(kind, {})["updated_at"] = datetime.utcnow().isoformat()
    META_FILE.write_text(json.dumps(meta, indent=2), encoding="utf-8")
    return get_prompt(kind)


def reset_prompt(kind: PromptKind) -> dict[str, Any]:
    return save_prompt(kind, _default_for(kind))


def load_template(kind: PromptKind) -> str:
    ensure_prompt_files()
    return _path_for(kind).read_text(encoding="utf-8")
