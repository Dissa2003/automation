# Chat Translator Test Automation

Automates testing of https://www.pixelssuite.com/chat-translator using Python Playwright.

Reads 50 Singlish test cases from the Excel file, submits each to the site, captures the Sinhala output, and writes **Actual Output** and **Status** (PASS/FAIL) back into the **same Excel file**.

## Project Structure

```
automation/
├── README.md
├── requirements.txt
├── data/
│   └── Assignment 1 - Test cases.xlsx   <- input and results stored here
└── test_automation/
    └── test_automation.py
```

## Setup

```bash
python -m pip install -r requirements.txt
playwright install chromium
```

## Run

```bash
python test_automation/test_automation.py
```

Results (Actual Output + PASS/FAIL) are written back into `data/Assignment 1 - Test cases.xlsx`.

## Options

| Flag | Default | Description |
|------|---------|-------------|
| `--excel PATH` | `data/Assignment 1 - Test cases.xlsx` | Source Excel file |
| `--url URL` | https://www.pixelssuite.com/chat-translator | Target URL |
| `--wait-ms N` | 8000 | Max ms to wait for API response per row |
| `--headless` | off | Run browser without UI |
| `--save-every N` | 0 | Save every N rows (0 = save at end only) |




git repo link-https://github.com/Dissa2003/automation.git