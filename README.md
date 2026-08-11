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
- **Automated Data Cleaning** — detects and reports duplicate rows, missing values, inconsistent text casing, and whitespace issues on load
- **Custom Cleaning via Natural Language** — describe a cleaning rule in plain English and the agent generates and applies the corresponding pandas code
- **Self-Correcting Agentic Loop** — code-generation tasks (Q&A, cleaning, SQL) follow an Analyze → Plan → Execute → Evaluate → Adapt cycle: if generated code fails, the error is fed back to the model, which retries with a corrected version (up to 3 attempts)
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
- **Automated EDA Reports** — key insights generated from summary statistics
- **Bilingual Output (Hinglish)** — reports and document Q&A responses are generated in Hinglish (Hindi + English, Roman script) for accessibility to the Indian developer/analyst market, while the codebase and all saved output files remain in professional English

## Tech Stack

- **Python 3.11+**
- **Google Gemini API** (`google-genai`) — natural language understanding and code generation
- **pandas** — data manipulation and cleaning
- **openpyxl** — Excel file generation, formulas, charts, conditional formatting, data validation
- **SQLite** — in-memory and persistent SQL query execution
- **python-docx**, **python-pptx**, **pypdf** — multi-format document parsing
- **Streamlit** — browser-based web interface

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
pip install streamlit
streamlit run app.py
```

Opens a browser-based UI with:
- Drag-and-drop file upload (CSV, Excel, PDF, Word, PowerPoint)
- One-click buttons for common actions (Report, Dashboard, Power BI export, SQL export, etc.)
- A chat box for free-form questions and command-prefixed instructions (`clean:`, `sql:`, `formula:`, `format:`, `validate:`)
- Download button for the most recently generated file

The web app reuses all the logic from `agent.py` directly — no duplicated code.


```bash
python agent.py your_data.csv
```

The agent also supports `.xlsx`, `.pdf`, `.docx`, and `.pptx` files.

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

<<<<<<< HEAD
- **Non-destructive by default** — operations write to new outputs instead of overwriting source data.
- **Formula-driven outputs** — Excel calculations use live formulas instead of hardcoded results where appropriate.
- **Centralized LLM wrapper** — model calls are routed through a single function to make provider changes easier.

## Limitations

- Native Power BI `.pbix` files cannot be generated directly; the agent produces clean exports and a DAX/dashboard guide.
- Native Excel PivotTable objects are not generated; equivalent summary tables are produced instead.
- Excel VBA/macros are not generated.
- The natural-language code execution path uses restricted `exec()` and is intended for personal/local use. A production multi-user deployment would require stronger sandboxing.
=======
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
>>>>>>> 58806ff (Add Streamlit interface and improve agent workflow)

## Development Note

<<<<<<< HEAD
This is an **AI-assisted learning project**. Claude was used extensively during development for code generation, debugging, refactoring, and implementation guidance. The project requirements, feature direction, testing, and final integration were reviewed and directed by the author.

The purpose of the project is to learn how LLMs can be integrated into practical data-analysis workflows rather than to present the implementation as entirely hand-written code.

## Future Improvements

- Web interface using Streamlit or Gradio
=======
>>>>>>> 58806ff (Add Streamlit interface and improve agent workflow)
- Multi-file joins and cross-file analysis
- Natural-language chart generation
- Write support for Word and PowerPoint files
- Stronger sandboxing for generated code

## Author

Sumeet Rawat
