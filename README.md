## Auto_gen_TC

### What this does

This repo contains small Python utilities to:

- Convert an Excel decision-tree template into JSON transitions (`excel_to_strings.py`)
- Generate **Test Cases (TCs)** as **all possible intent paths** from a root node (`tcgen.py`)

### Install (needed for Excel conversion)

`excel_to_strings.py` requires `pandas` (and an Excel engine such as `openpyxl`).

```bash
python3 -m pip install -r requirements.txt
```

### Input format (for `tcgen.py`)

`tcgen.py` expects a JSON list where each item is a transition:

```json
[
  {
    "step_no": "A1",
    "step_name": "Chào và giới thiệu khoản vay",
    "customer_intent": "KH không đồng ý",
    "next_step": "A2",
    "action_code": "No_need"
  }
]
```

Special rule implemented:

- If multiple rows share the same `(step_no, next_step)`, their `customer_intent` values are **merged** into one grouped intent like:  
  `Intent 1 / Intent 2 / Intent 3`

### Convert Excel to JSON transitions

If you have an Excel file in `./input/`, you can convert it into the JSON transition list used by `tcgen.py`.

Auto-pick the first Excel file in `./input/` and write JSON to `output/rows.json`:

```bash
python3 excel_to_strings.py --out output/rows.json
```

Specify an Excel file explicitly:

```bash
python3 excel_to_strings.py --file "input/(FEC_VoiceFlow) KịchBản_XS_2026.xlsx" --out output/rows.json
```

If your workbook has multiple sheets, you can set a sheet name or index:

```bash
python3 excel_to_strings.py --file "input/(FEC_VoiceFlow) KịchBản_XS_2026.xlsx" --sheet 0 --out output/rows.json
```

### One-command pipelines

Run the full flow in one command:

- `pipeline_e2e.py`: `excel_to_json.py` -> `tcgen_e2e_human.py` -> `tc_to_excel.py`
- `pipeline_e2e_max.py`: `excel_to_json.py` -> `tcgen_e2e_max.py` -> `tc_to_excel.py` (same as E2E human, but treats `lần N` repeat intents as normal transitions)
- `pipeline_e2e_short.py`: `excel_to_json.py` -> `tcgen_e2e_short.py` -> `tc_to_excel.py`
- `pipeline_output_short.py`: `excel_to_json.py` -> `tcgen_output_short.py` -> `tc_to_excel.py`
- `pipeline_output.py`: `excel_to_json.py` -> `tcgen_output_human.py` -> `tc_to_excel.py`
- `pipeline_multi_responses.py`: `excel_to_json.py` -> `tcgen_multi_responses.py` -> `tc_to_excel.py`
- `pipeline_all.py`: runs `e2e_short`, `output_short`, and `multi_responses`, then exports all sheets into one Excel workbook.

Examples:

```bash
python3 pipeline_e2e.py --file "input/FEC_VoiceAgent_BRD.xlsx"
python3 pipeline_e2e_max.py --file "input/FEC_VoiceAgent_BRD.xlsx"
python3 pipeline_e2e_short.py --file "input/FEC_VoiceAgent_BRD.xlsx"
python3 pipeline_output_short.py --file "input/FEC_VoiceAgent_BRD.xlsx"
python3 pipeline_output.py --file "input/FEC_VoiceAgent_BRD.xlsx"
python3 pipeline_multi_responses.py --file "input/VB_M1_Revamp_Credit_card.xlsx"
python3 pipeline_all.py --file "input/VB_M1_Revamp_Credit_card.xlsx" --no-gen-data
```

If `--file` is omitted, the pipeline auto-picks the first non-temporary Excel file in `./input/`.

Default outputs:

python3 pipeline_multi_responses.py \
  --file "input/VB_M1_Revamp_Credit_card.xlsx" \
  --rows-out output/rows_multi_vb.json \
  --testcases-out output/testcases_multi_vb.json \
  --excel-out output/testcases_multi_vb.xlsx

python3 pipeline_e2e.py \
  --file "input/VB_M1_Revamp_Credit_card.xlsx" \
  --rows-out output/rows_e2e_vb.json \
  --testcases-out output/testcases_e2e_vb.json \
  --excel-out output/testcases_e2e_vb.xlsx

python3 pipeline_e2e_short.py \
  --file "input/VB_M1_Revamp_Credit_card.xlsx" \
  --rows-out output/rows_e2e_short_vb.json \
  --testcases-out output/testcases_e2e_short_vb.json \
  --excel-out output/testcases_e2e_short_vb.xlsx

- `pipeline_e2e.py`
  - `output/rows_e2e.json`
  - `output/testcases_e2e.json`
  - `output/testcases_e2e.xlsx`
- `pipeline_e2e_max.py`
  - `output/rows_e2e_max.json`
  - `output/testcases_e2e_max.json`
  - `output/testcases_e2e_max.xlsx`
- `pipeline_e2e_short.py`
  - `output/rows_e2e_short.json`
  - `output/testcases_e2e_short.json`
  - `output/testcases_e2e_short.xlsx`
- `pipeline_output_short.py`
  - `output/rows_output_short.json`
  - `output/testcases_output_short.json`
  - `output/testcases_output_short.xlsx`
- `pipeline_output.py`
  - `output/rows_output.json`
  - `output/testcases_output.json`
  - `output/testcases_output.xlsx`
- `pipeline_multi_responses.py`
  - `output/rows_multi_responses.json`
  - `output/testcases_multi_responses.json`
  - `output/testcases_multi_responses.xlsx`
- `pipeline_all.py`
  - `output/rows_e2e_short.json`
  - `output/testcases_e2e_short.json`
  - `output/rows_output_short.json`
  - `output/testcases_output_short.json`
  - `output/rows_multi_responses.json`
  - `output/testcases_multi_responses.json`
  - `output/testcases_all.xlsx`

Run all pipelines into one workbook:

```bash
python3 pipeline_all.py \
  --file "input/VB_M1_Revamp_Credit_card.xlsx" \
  --suffix vb \
  --no-gen-data
```

Useful options:

- `--file`: input Excel file. If omitted, auto-picks the first non-temporary Excel file in `./input/`.
- `--sheet`: sheet name or index. If omitted, auto-detects the template sheet/header row.
- `--root`: root step. If omitted, auto-detects from the sheet.
- `--max-depth`: traversal limit, default `200`.
- `--suffix`: suffix for output filenames, e.g. `--suffix vb` writes `output/testcases_all_vb.xlsx`.
- `--no-gen-data`: skip hybrid test-data generation. This is faster and recommended for dry runs.

You can override any output path:

```bash
python3 pipeline_e2e.py \
  --file "input/FEC_VoiceAgent_BRD.xlsx" \
  --rows-out output/rows2.json \
  --testcases-out output/testcasesv2.json \
  --excel-out output/testcasesv2.xlsx
```

### Generate test cases

From the workspace root:

```bash
python3 tcgen.py --in output/rows.json --root A1 --out output/testcases.json --format json
```

CSV output:

```bash
python3 tcgen.py --in output/rows.json --root A1 --out output/testcases.csv --format csv
```

Output shape (JSON):

```json
[
  { "tc_id": "TC1", "steps": ["intent group at level 1", "intent group at level 2"] }
]
```

Notes:

- Terminal nodes are detected when `next_step` is `End` (case-insensitive) or missing/blank.
- Cycles like `A1 -> A1` are prevented from infinite traversal; the TC ends with a `(cycle to <node>)` marker.

### Export test cases to Excel (`tc_to_excel.py`)

Convert the TC JSON to Excel with columns: **TC_ID**, **Test Scenario**, **Expected Action Code**, **Test Data**.

```bash
python3 tc_to_excel.py --in output/testcases_v2.json --out output/test_cases.xlsx
```

### Generate test data with Llama 3.1 (`generate_test_data.py`)

Fill the **Test Data** column with realistic Vietnamese customer utterances using Ollama (Llama 3.1).

**Requirements:** [Ollama](https://ollama.com) installed and running locally, with a model pulled (e.g. `ollama pull llama3.1`). Install the Python client: `pip install ollama`.

From **JSON** (tcgen output) → Excel with test data:

```bash
python3 generate_test_data.py --in output/testcases_v2.json --out output/test_cases.xlsx --model llama3.1
```

From **existing Excel** (fill Test Data column in place or to a new file):

```bash
python3 generate_test_data.py --in output/test_cases.xlsx --out output/test_cases.xlsx --model llama3.1
```

Optional: `--limit N` to process only the first N test cases (for testing).

### Generate test data with tls_client_bot (`generate_test_data_nlp.py`)
Fill the **Test Data** column with Vietnamese customer utterances using the local `tls_client_bot` NLP intent classifier (instead of Llama).

From **JSON** (tcgen output) → Excel with test data:

```bash
python3 generate_test_data_neuron.py --in output/testcases_v2.json --out output/test_cases.xlsx
```

From **existing Excel** (fill Test Data column):

```bash
python3 generate_test_data_neuron.py --in output/test_cases.xlsx --out output/test_cases.xlsx
```

### Train bot
```bash
 source venv/bin/activate
 pip install Flask torch torchvision nltk
 python3 train.py
 ```

 ### exel intents file to intent json file
 ```bash
 python3 tls_client_bot/excel_to_intents_json.py \
  -i tls_client_bot/data_excel/voice_flow_data_input.xlsx \
  -o tls_client_bot/intents_from_excel.json 
  ```
