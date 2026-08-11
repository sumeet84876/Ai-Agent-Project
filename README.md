# AI Data Analyst Agent

An AI-powered command-line agent that reads, cleans, analyzes, and reports on data from CSV, Excel, PDF, Word, and PowerPoint files using natural language and Google's Gemini API.

## Overview

This project explores how an AI assistant can automate repetitive parts of a data analyst workflow, including data cleaning, exploratory analysis, natural-language questions, SQL generation, Excel dashboard creation, and Power BI-ready exports.

## Features

- **Universal File Reader** — supports `.csv`, `.xlsx`, `.pdf`, `.docx`, `.pptx`
- **Automated Data Cleaning** — detects duplicate rows, missing values, inconsistent text casing, and whitespace issues
- **Natural-Language Cleaning** — converts plain-English cleaning instructions into pandas operations
- **Natural-Language Q&A** — ask questions about loaded data or documents
- **SQL Query Generation** — converts natural-language questions into SQL queries and executes them against SQLite
- **Persistent SQLite Export** — saves datasets as `.db` files for use with SQL clients
- **Excel Dashboard Generation** — creates formula-driven workbooks with charts
- **Advanced Excel Features** — formula columns, conditional formatting, data validation, and in-place workbook editing
- **Power BI Integration** — exports clean data and generates a guide with DAX measures and dashboard-building steps
- **Automated EDA Reports** — generates summaries and insights from descriptive statistics

## Tech Stack

- Python 3.11+
- Google Gemini API (`google-genai`)
- pandas
- openpyxl
- SQLite
- python-docx
- python-pptx
- pypdf
- Streamlit

## Setup

### 1. Get a Gemini API key

Create a Gemini API key through Google AI Studio.

### 2. Install dependencies

```bash
pip install -r requirements.txt
```

### 3. Set your API key

Windows PowerShell:

```powershell
$env:GEMINI_API_KEY="your-key-here"
```

macOS/Linux:

```bash
export GEMINI_API_KEY="your-key-here"
```

## Usage

### Command-line interface

```bash
python agent.py your_data.csv        # also supports .xlsx, .pdf, .docx, .pptx
```

### Web interface (Streamlit)

```bash
streamlit run app.py
```

Opens a browser-based UI with:
- Drag-and-drop file upload (CSV, Excel, PDF, Word, PowerPoint)
- One-click buttons for common actions (Report, Dashboard, Power BI export, SQL export, etc.)
- A chat box for free-form questions and command-prefixed instructions (`clean:`, `sql:`, `formula:`, `format:`, `validate:`)
- Download button for the most recently generated file

The web app reuses all the logic from `agent.py` directly — no duplicated code.

## Commands

| Command | Description |
|---|---|
| `report` | Generate an automated EDA summary |
| `sql: <question>` | Generate and execute a SQL query |
| `clean: <instruction>` | Apply a natural-language cleaning instruction |
| `save` | Export the cleaned dataset to CSV and Excel |
| `sql-export` | Save the dataset as a SQLite database |
| `dashboard` | Generate an Excel dashboard |
| `powerbi` | Export Power BI-ready data |
| `powerbi-guide` | Generate DAX measures and dashboard instructions |
| `formula: <instruction>` | Add a formula-driven Excel column |
| `format: <instruction>` | Apply Excel conditional formatting |
| `validate: <instruction>` | Add Excel data validation |
| `edit-excel` | Edit an existing Excel workbook in place |
| `<any question>` | Ask a natural-language question about the data |
| `exit` | Exit the agent |

## Design Notes

- **Non-destructive by default** — every operation writes to a new file rather than overwriting the source, so the original data is never at risk.
- **Formulas over hardcoded values** — Excel outputs (dashboards, formula columns) use live formulas so they stay correct if the underlying data changes.
- **Provider-agnostic LLM wrapper** — all LLM calls go through a single `call_llm()` function, making it straightforward to swap providers if needed.
- **Self-correcting execution** — `generate_and_execute_with_retry()` implements a generic Analyze → Plan → Execute → Evaluate → Adapt loop used across Q&A, cleaning, and SQL generation, so transient code errors are automatically corrected rather than surfaced directly to the user.

## Limitations

- Cannot generate native Power BI `.pbix` files directly — Power BI Desktop has no public API for this. The agent instead produces clean data exports and a DAX/dashboard-building guide.
- Cannot generate true native Excel PivotTable objects (limited library support); produces equivalent summary tables instead.
- Does not support Excel macros/VBA generation.
- `exec()` is used with restricted builtins for natural-language-to-code execution; this is adequate for personal/local use but would need additional sandboxing (e.g. subprocess isolation) for a multi-user or production deployment.
- The self-correction loop retries up to 3 times; persistent failures (e.g. a genuinely unanswerable question) still surface as an error after the final attempt.

## Development Note

This is an **AI-assisted learning project**. Claude was used extensively during development for code generation, debugging, refactoring, and implementation guidance. The project requirements, feature direction, testing, and final integration were reviewed and directed by the author.

The purpose of the project is to learn how LLMs can be integrated into practical data-analysis workflows rather than to present the implementation as entirely hand-written code.

## Future Improvements

- Multi-file joins and cross-file analysis
- Natural-language chart generation
- Write support for Word and PowerPoint files
- Stronger sandboxing for generated code

## Author

Sumeet Rawat
