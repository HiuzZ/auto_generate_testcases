#!/usr/bin/env python3
import argparse
import json
from openpyxl import Workbook

def json_to_excel_flat(json_path, excel_path, sheet_name="Sheet1"):
    """
    Convert intents JSON to a flat Excel file.
    Each intent -> one row:
      Col A: tag
      Col B: patterns combined as "1. pat1\n2. pat2..."
      Col C: responses combined as "1. resp1\n2. resp2..."
    """
    with open(json_path, "r", encoding="utf-8") as f:
        data = json.load(f)

    intents = data.get("intents", [])
    wb = Workbook()
    ws = wb.active
    ws.title = sheet_name

    # Header
    ws.append(["Tag", "Patterns", "Responses"])

    for intent in intents:
        tag = intent.get("tag", "")
        patterns = intent.get("patterns", [])
        responses = intent.get("responses", [])

        # Combine patterns into a numbered multiline string
        patterns_text = "\n".join(
            f"{i+1}. {p}" for i, p in enumerate(patterns)
        )

        # Combine responses into a numbered multiline string
        responses_text = "\n".join(
            f"{i+1}. {r}" for i, r in enumerate(responses)
        )

        ws.append([tag, patterns_text, responses_text])

    wb.save(excel_path)
    print(f"Excel file saved: {excel_path}")


def main():
    parser = argparse.ArgumentParser(
        description="Convert intents JSON to Excel (1 row per intent, patterns & responses as numbered lists)."
    )
    parser.add_argument("--input", "-i", required=True, help="Input JSON file")
    parser.add_argument("--output", "-o", default="intents_flat.xlsx", help="Output Excel file")
    parser.add_argument("--sheet", "-s", default="Sheet1", help="Sheet name")
    args = parser.parse_args()

    json_to_excel_flat(args.input, args.output, args.sheet)


if __name__ == "__main__":
    main()