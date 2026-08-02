# AI Data Analyst Agent

An AI-powered command-line agent that reads, cleans, analyzes, and reports on data from CSV, Excel, PDF, Word, and PowerPoint files using natural language — powered by Google's Gemini API (free tier).

## Overview

This project automates the repetitive parts of a data analyst's workflow: cleaning messy datasets, answering ad-hoc questions in plain English, generating SQL queries, building Excel dashboards, and preparing Power BI-ready exports — all through a simple conversational interface.

## Features

- **Universal File Reader** — supports `.csv`, `.xlsx`, `.pdf`, `.docx`, `.pptx`
- **Automated Data Cleaning** — detects and reports duplicate rows, missing values, inconsistent text casing, and whitespace issues on load
- **Custom Cleaning via Natural Language** — describe a cleaning rule in plain English and the agent generates and applies the corresponding pandas code
- **Natural Language Q&A** — ask questions about your data or documents and get direct answers, no code required
- **SQL Query Generation** — describe what you want in English, get a working SQL query and its result
- **Persistent SQLite Export** — save your dataset as a real `.db` file that opens in any SQL client (SQL Workbench, DB Browser for SQLite, etc.)
- **Excel Dashboard Generation** — creates a multi-sheet workbook with live formulas (SUM, AVERAGE, MAX, MIN) and charts, not hardcoded values
- **Advanced Excel Features**:
  - Formula-driven columns (nested IF, INDEX-MATCH, VLOOKUP)
  - Conditional formatting (cell rules, color scales)
  - Data validation (dropdown lists)
  - In-place editing of existing Excel files that preserves other sheets/tabs
- **Power BI Integration** — exports clean, typed data ready for import, plus a generated guide with DAX measures and dashboard-building steps
- **Bilingual Output (Hinglish)** — reports, guides, and document Q&A responses are generated in Hinglish (Hindi + English, Roman script) for accessibility to the Indian developer/analyst market, while the codebase itself is fully in English
- **Automated EDA Reports** — key insights generated from summary statistics

## Tech Stack

- **Python 3.11+**
- **Google Gemini API** (`google-genai`) — natural language understanding and code generation
- **pandas** — data manipulation and cleaning
- **openpyxl** — Excel file generation, formulas, charts, conditional formatting, data validation
- **SQLite** — in-memory and persistent SQL query execution
- **python-docx**, **python-pptx**, **pypdf** — multi-format document parsing

## Setup

### 1. Get a free Gemini API key
Visit [aistudio.google.com/apikey](https://aistudio.google.com/apikey), sign in with a Google account, and generate a key. No credit card required.

### 2. Install dependencies
```bash
pip install google-genai pandas openpyxl python-docx python-pptx pypdf
```

### 3. Set your API key
```bash
export GEMINI_API_KEY="your-key-here"        # macOS/Linux
$env:GEMINI_API_KEY="your-key-here"          # Windows PowerShell
```

## Usage

```bash
python agent.py your_data.csv        # also supports .xlsx, .pdf, .docx, .pptx
```

### Commands (tabular files)

| Command | Description |
|---|---|
| `report` | Generate an automated EDA summary with key insights |
| `sql: <question>` | Convert a natural language question into a SQL query and run it |
| `clean: <instruction>` | Apply a custom cleaning rule described in plain English |
| `save` | Export the current (cleaned) dataset to CSV and Excel |
| `sql-export` | Save the dataset as a persistent `.db` file for external SQL tools |
| `dashboard` | Generate an Excel dashboard with live formulas and charts |
| `powerbi` | Export a clean, Power BI import-ready Excel file |
| `powerbi-guide` | Generate DAX measures and a step-by-step Power BI dashboard guide |
| `formula: <instruction>` | Add a new column driven by a live Excel formula (IF, INDEX-MATCH, etc.) |
| `format: <instruction>` | Apply conditional formatting to a column |
| `validate: <instruction>` | Add a dropdown (data validation) to a column |
| `edit-excel` | Update an existing Excel file in place while preserving other sheets |
| `<any question>` | Ask anything about the data in plain English |
| `exit` | Quit the agent |

### Example session

```
$ python agent.py sales_data.csv
Loading sales_data.csv...
Loaded: 1020 rows, 10 columns

Cleaning log:
  - Removed 3 duplicate rows
  - Missing values found: Salary: 24, Age: 211
  - 'Region' has inconsistent casing — use 'clean:' to normalize

You: clean: normalize Region to proper case
You: report
You: sql: total sales by region, highest to lowest
You: dashboard
You: exit
```

For document files (PDF/Word/PPT), the agent switches to a simple Q&A mode:
```
$ python agent.py contract.pdf
Document loaded (18,432 characters)
You: summarize the key terms of this document
You: what is the termination clause?
```

## Design Notes

- **Non-destructive by default** — every operation writes to a new file rather than overwriting the source, so the original data is never at risk.
- **Formulas over hardcoded values** — Excel outputs (dashboards, formula columns) use live formulas so they stay correct if the underlying data changes.
- **Provider-agnostic LLM wrapper** — all LLM calls go through a single `call_llm()` function, making it straightforward to swap providers if needed.

## Limitations

- Cannot generate native Power BI `.pbix` files directly — Power BI Desktop has no public API for this. The agent instead produces clean data exports and a DAX/dashboard-building guide.
- Cannot generate true native Excel PivotTable objects (limited library support); produces equivalent summary tables instead.
- Does not support Excel macros/VBA generation.
- `exec()` is used with restricted builtins for natural-language-to-code execution; this is adequate for personal/local use but would need additional sandboxing (e.g. subprocess isolation) for a multi-user or production deployment.

## Possible Extensions

- Web interface via Streamlit or Gradio
- Multi-file joins and cross-file analysis
- Chart generation from natural language requests
- Write support for Word/PowerPoint files (currently read-only)

## Author

Built by Sumeet as a personal project to explore practical applications of LLMs in data analytics workflows.