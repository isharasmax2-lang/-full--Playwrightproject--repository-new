# Playwright Test Automation — SwiftTranslator

This repository contains the Playwright-based test automation used to exercise the SwiftTranslator frontend.

Contents
- `test_automation.py` — main test runner (Playwright + openpyxl)
- `Assignment 1 - Test cases.xlsx` — completed test cases workbook (if present in repository root)
- `requirements.txt` — Python dependencies

Quick start (Windows)

1. Create and activate a virtual environment:

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
```

2. Install dependencies:

```powershell
pip install -r requirements.txt
python -m playwright install
```

3. Run the tests (example):

```powershell
python test_automation.py --excel "Assignment 1 - Test cases.xlsx" --headless
```

Notes
- The script writes results back to the Excel file and produces an HTML report and per-test screenshots under `results/`.
- If you want to watch the browser, omit `--headless`.
- You can customize `--results-dir`, `--sheet`, and other flags. See `python test_automation.py --help`.

How to publish to GitHub (optional)

If you want me to push this repository to GitHub, provide a personal access token (PAT) or authorize via `gh auth login`. Alternatively, run:

```powershell
git remote add origin https://github.com/your-username/your-repo.git
git branch -M main
git push -u origin main
```

License: provided as-is for submission purposes.
