# Excel to simPRO Quote Tool

This repository holds the code for the Excel → simPRO quoting workflow.

## Purpose

- Read quote/job details from the Excel pricing workbook
- Help search work tasks using plain-English terms
- Prepare quote lines for sending into simPRO
- Keep the code in GitHub so updates are easier and safer

## Current status

Starter repo created from ChatGPT. The current scripts from the Mac still need to be uploaded if we want to preserve the existing live version.

## Suggested files to upload from the Mac

- `excel_to_simpro_quote_with_lines.py`
- `task_search.py`
- `send_quote_to_simpro.command`
- any exported Excel VBA modules
- sample task/pricing data with sensitive/customer data removed

## Quick local run

```bash
python3 task_search.py
```
