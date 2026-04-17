# 🎙️ Hybrid Test Data Generator cho Chatbot

## 🧠 Tổng quan

Công cụ này tự động sinh **dữ liệu kiểm thử (test data)** dạng hội thoại thực tế từ các test case đã được định nghĩa (dạng JSON hoặc Excel).  
Điểm đặc biệt: sử dụng **chiến lược hybrid** – ưu tiên dùng **TLS Client Bot** (NLP intent classifier) để sinh lời thoại khách hàng; nếu TLS bot không xử lý được một bước nào đó, nó sẽ **fallback sang Llama 3.1** (qua Ollama) chỉ cho riêng bước đó.

Công cụ này còn hỗ trợ:
- **Multi‑intent steps**: Một bước có thể chứa nhiều ý định, phân cách bằng ` \ `.  
- **Multiple rounds**: Với mỗi test case, sinh nhiều round (vòng) – round thứ `k` chọn intent thứ `k` từ mỗi step (xoay vòng tròn).  
- **Bot response variants**: Nếu bot response có nhiều biến thể (cũng phân cách ` \ `), chọn ngẫu nhiên một biến thể cho mỗi lần sinh.  
- Đầu ra là file Excel với các cột: `TC_ID`, `Round`, `Test Scenario`, `Bot Responses`, `Path`, `Expected Action Code`, `Test Data`.

---

## 🧩 Các thành phần chính

| Thành phần | Mô tả |
|------------|-------|
| **TLS Client Bot** | Hệ thống NLP dựa trên intents.json, có khả năng phân loại ý định và sinh câu trả lời. Được ưu tiên sử dụng. |
| **Llama 3.1 (Ollama)** | Mô hình ngôn ngữ lớn, dùng làm fallback khi TLS bot không sinh được câu cho một bước cụ thể. |
| **Hybrid per‑step fallback** | Không fallback toàn bộ test case, mà chỉ fallback từng bước thất bại, giữ nguyên các bước thành công từ TLS bot. |
| **Multi‑intent parser** | Tách chuỗi step dạng `"intent1 \ intent2 \ intent3"` thành danh sách các intent riêng lẻ. |
| **Circular round selection** | Với mỗi test case, xác định số round = độ dài lớn nhất của danh sách intent trong các step. Round `r` chọn intent thứ `(r-1) % len(intents)` từ mỗi step. |
| **Random bot response variant** | Bot response có thể có nhiều biến thể (vd: `"cảm ơn \ cám ơn"`), mỗi lần sinh chọn ngẫu nhiên một biến thể. |
| **Sanitization & normalization** | Làm sạch câu sinh ra: viết thường, loại bỏ dấu câu cuối, bỏ ký tự đặc biệt, giới hạn số từ (mặc định 12 từ). |

---

## ⚙️ Luồng xử lý chính

```mermaid
graph TD
    A[Đọc input: JSON hoặc Excel] --> B[Phân tích test case]
    B --> C[Mỗi test case: tách steps thành list intents bằng dấu \]
    C --> D[Tính max_rounds = max(len(intents) per step)]
    D --> E[Với mỗi round r từ 1..max_rounds]
    E --> F[Chọn intent thứ r cho mỗi step (circular)]
    F --> G[Gọi hybrid generation cho round này]
    G --> H[TLS bot thử sinh cho từng step]
    H --> I{Thành công?}
    I -->|Có| J[Ghi nhận câu]
    I -->|Không| K[Đánh dấu step cần fallback]
    K --> L[Với mỗi step fallback: dùng Llama với context từ các bước trước]
    L --> M[Kết hợp tất cả câu thành Test Data]
    J --> M
    M --> N[Ghi một dòng vào Excel (TC_ID, Round, Test Scenario, Bot Responses, Path, Expected Action Code, Test Data)]
    N --> E
```

### Giải thích các bước

1. **Đọc input**  
   - Nếu input là JSON (định dạng từ `tcgen`): mỗi object có `steps`, `bot_responses`, `expected_action_code`, `path`.  
   - Nếu input là Excel: đọc cột `Test Scenario` (chứa các bước đánh số), `Bot Responses`, v.v. Tái cấu trúc lại thành list steps.

2. **Phân tích multi‑intent**  
   Mỗi step (chuỗi) được tách bằng regex `\s*\\\s*` (dấu `\` có khoảng trắng tùy ý).  
   Ví dụ: `"Kiểm tra số dư \ Hỏi lãi suất"` → `["Kiểm tra số dư", "Hỏi lãi suất"]`.

3. **Xác định số round**  
   `max_rounds = max(len(intents) for intents in steps_intents)`.  
   Nếu step nào có ít intent hơn, sẽ được lặp lại theo vòng tròn.

4. **Sinh hybrid cho một round**  
   - **Bước 1 – TLS bot cho tất cả step**:  
     Gọi `_generate_step_with_tls_bot(step)`. Hàm này chuẩn hóa step thành message (giữ lại attempt number nếu có), gọi `tls_client_bot.get_response()`. Nếu không có hoặc lỗi, dùng fallback matching từ `intents.json`.  
   - **Bước 2 – Xác định step thất bại**:  
     Những step trả về `"(không sinh được)"` được đánh dấu.  
   - **Bước 3 – Llama fallback cho từng step thất bại**:  
     Với mỗi step thất bại, gọi `_generate_single_with_llama()`.  
     Prompt được xây dựng có bao gồm **context** từ các câu đã sinh thành công ở các bước trước đó (cùng round), giúp Llama hiểu được diễn biến hội thoại.  
     Ngoài ra, nếu bot response có nhiều biến thể, chọn ngẫu nhiên một biến thể để đưa vào prompt (giúp Llama sinh câu phù hợp).  
   - **Bước 4 – Ghép kết quả**:  
     Các câu được nối với nhau bằng dấu xuống dòng và đánh số thứ tự bước.

5. **Ghi vào Excel**  
   Mỗi round là một dòng riêng, với cột `Round` chỉ rõ vòng thứ mấy.  
   Cột `Test Data` chứa các câu thoại đã sinh (theo thứ tự bước).  
   Tự động điều chỉnh độ rộng cột và wrap text.

---

## 🔧 Cơ chế hybrid chi tiết

### 1. TLS Client Bot
- **Yêu cầu**: Thư mục `tls_client_bot` phải chứa `intents.json` và code `chat.py` có hàm `get_response()`.  
- **Cách hoạt động**:  
  - Hàm `_normalize_step_to_bot_message` biến đổi step (có thể chứa dấu `\` và attempt number) thành câu hỏi phù hợp để gửi đến bot.  
    - Với multi‑intent, chỉ lấy intent đầu tiên, nhưng giữ lại attempt number cao nhất (vd: `"hỏi lãi suất lần 2"`).  
    - Với single intent, cố gắng match chính xác với tag trong intents.json.  
  - Gọi `get_response(message)` – nếu thành công trả về câu trả lời.  
  - Nếu bot không hiểu (chuỗi chứa `"do not understand"`) hoặc exception, chuyển sang fallback.  
- **Fallback khi không có TLS bot**:  
  Dùng chính intents.json để tìm pattern matching (so khớp sau khi bỏ dấu, chuẩn hóa Unicode). Nếu tìm thấy tag, chọn ngẫu nhiên một response từ danh sách `responses` của tag đó.

### 2. Llama 3.1 (Ollama)
- **Kích hoạt**: Chỉ khi `ollama` được import thành công và có model `llama3.1`.  
- **Prompt engineering**:  
  - Yêu cầu LLM chỉ sinh **một câu duy nhất** cho bước hiện tại.  
  - Cung cấp context các câu trước đó (từ các bước đã xử lý thành công).  
  - Nếu có `expected_bot_response` (được chọn ngẫu nhiên từ variants), đưa vào prompt để định hướng.  
  - Ràng buộc nghiêm ngặt: viết thường, không dấu câu cuối, tối đa 12 từ, không ký tự đặc biệt.  
- **Xử lý output**:  
  - Loại bỏ số thứ tự nếu LLM vô tình thêm.  
  - Gọi `_sanitize_utterance` để làm sạch.  

### 3. Per‑step context
Điểm mạnh: Khi fallback cho step thứ `i`, context bao gồm các câu đã sinh thành công từ step `1..i-1`.  
Điều này giúp Llama hiểu được mạch hội thoại, tránh sinh câu lạc lõng hoặc lặp lại thông tin.

Ví dụ:
- Step 1 (TLS thành công): `"tôi muốn kiểm tra số dư"`  
- Step 2 (TLS thất bại, fallback Llama): context = `["tôi muốn kiểm tra số dư"]`, prompt yêu cầu sinh câu cho `"hỏi lãi suất lần 1"`.  
  Llama có thể sinh: `"lãi suất cho vay hiện tại là bao nhiêu phần trăm"` (tự nhiên, không bị lặp).

---

## 📐 Xử lý multi‑intent và multiple rounds

### Định dạng step
Một step có thể chứa nhiều ý định, viết liền nhau, phân cách bằng ` \ ` (backslash có khoảng trắng hai bên).  
Ví dụ:  
```
Kiểm tra số dư \ Hỏi lãi suất lần 1 \ Hỏi lãi suất lần 2
```

### Cách sinh các round
- **max_rounds** = độ dài lớn nhất của danh sách intent trong các step.  
- Với round `r` (bắt đầu từ 1), intent được chọn cho step `j` là `intents[j][(r-1) % len(intents[j])]`.  

Ví dụ:
- Step 1: `["A", "B", "C"]` (3 intents)  
- Step 2: `["X", "Y"]` (2 intents)  
- Step 3: `["P"]` (1 intent)  

→ max_rounds = 3  

| Round | Step 1 | Step 2 | Step 3 |
|-------|--------|--------|--------|
| 1     | A      | X      | P      |
| 2     | B      | Y      | P      |
| 3     | C      | X      | P      |

### Bot response variants
Tương tự, bot response có thể có nhiều biến thể phân cách bằng ` \ `. Mỗi lần sinh test data cho một round, chọn **ngẫu nhiên** một biến thể cho mỗi bước (độc lập).  
Biến thể này được dùng để:
- Đưa vào prompt Llama (giúp LLM sinh câu phù hợp với phản hồi mong đợi).
- Ghi vào cột `Bot Responses` (để người đánh giá biết bot kỳ vọng trả lời thế nào).

---

## 🧪 Các hàm quan trọng

### `_generate_hybrid_for_round(steps, bot_responses, tc_id, round_num)`
- **Đầu vào**: danh sách các step (đã chọn intent cụ thể cho round này), danh sách bot responses gốc, id test case, số round.  
- **Đầu ra**: chuỗi test data, mỗi bước trên một dòng, đánh số.  
- **Quy trình**:  
  1. Thử TLS bot cho tất cả step, lưu kết quả và đánh dấu step thất bại.  
  2. Với mỗi step thất bại, gọi Llama với context từ các step đã thành công trước đó.  
  3. Ghép kết quả.

### `_normalize_step_to_bot_message(step)`
Biến đổi step (có thể multi‑intent, có attempt number) thành câu gửi đến TLS bot.  
- Nếu step có `"im lặng"` hoặc `"không nghe rõ"`, trả về tag tương ứng.  
- Multi‑intent: lấy intent đầu tiên, giữ lại attempt number cao nhất (vd: `"hỏi lãi suất lần 2"`).  
- Single intent: cố gắng match chính xác với tag trong intents.json.

### `_generate_single_with_llama(step, step_number, total_steps, context, expected_bot_response)`
Xây dựng prompt chi tiết, gọi Ollama, làm sạch kết quả.  
Prompt có dạng:
```
Bạn là KHÁCH HÀNG ... Nhiệm vụ: chỉ tạo 1 câu lời thoại KH cho bước số X.
Các câu trước đó: ...
Kịch bản bước X: ...
Bot sẽ đáp: "..."
Yêu cầu: ... (viết thường, không dấu câu, tối đa 12 từ)
Đầu ra:
```

### `_sanitize_utterance(text)`
- Chuyển thành chữ thường.  
- Xóa dấu phẩy, ký tự đặc biệt `*^$#@%&?!`.  
- Xóa dấu chấm, chấm hỏi, chấm than ở cuối.  
- Giới hạn số từ (mặc định 12).  
- Nếu rỗng hoặc `"do not understand"`, trả về `"(không sinh được)"`.

---

## 📁 Đầu ra Excel

| Cột | Ý nghĩa | Ví dụ |
|-----|---------|-------|
| `TC_ID` | Mã test case | `TC001` |
| `Round` | Vòng thứ mấy | `1`, `2`, `3` |
| `Test Scenario` | Các bước hội thoại (đánh số, mỗi bước là intent đã chọn) | `1. Kiểm tra số dư`<br>`2. Hỏi lãi suất lần 1` |
| `Bot Responses` | Các phản hồi mong đợi của bot (đánh số, mỗi bước một response, đã chọn variant ngẫu nhiên) | `1. Số dư của bạn là 1 triệu`<br>`2. Lãi suất hiện tại 5%` |
| `Path` | Chuỗi các node đã qua | `A1 -> B2 -> End` |
| `Expected Action Code` | Mã hành động kỳ vọng | `INQUIRY` |
| `Test Data` | **Câu thoại thực tế của khách hàng** sinh ra (đánh số bước) | `1. cho em hỏi số dư tài khoản`<br>`2. lãi suất vay bây giờ là bao nhiêu ạ` |

File Excel được tự động căn chỉnh độ rộng cột và wrap text cho dễ đọc.

---

## 💡 Ưu điểm của hybrid per‑step approach

1. **Tận dụng tốc độ và tính ổn định của rule‑based / intent‑based bot** cho các bước đơn giản.  
2. **Dùng LLM chỉ khi cần** – tiết kiệm tài nguyên và thời gian so với gọi LLM cho toàn bộ test case.  
3. **Giữ nguyên context** – Llama được cung cấp các câu trước đó (từ TLS bot), giúp câu sinh ra tự nhiên và mạch lạc.  
4. **Xử lý multi‑intent linh hoạt** – sinh đủ các tổ hợp intent mà không cần nhân bản test case thủ công.  
5. **Ngẫu nhiên hóa bot response variants** – tăng tính đa dạng cho dữ liệu kiểm thử.

---

## 🔗 Phụ thuộc

- **openpyxl** – đọc/ghi Excel.  
- **ollama** (tuỳ chọn) – nếu muốn dùng Llama fallback.  
- **tls_client_bot** – cần có thư mục chứa `intents.json` và module `chat.py`.  
- **unicodedata** – chuẩn hóa dấu tiếng Việt.  
- **random, re, json, pathlib** – thư viện chuẩn.

Công cụ này đặc biệt hữu ích khi bạn có sẵn một tập các test case dạng kịch bản (được sinh từ decision tree) và muốn tự động tạo ra các hội thoại thực tế để chạy kiểm thử end‑to‑end.