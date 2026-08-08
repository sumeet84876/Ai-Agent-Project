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

- **Non-destructive by default** — operations write to new outputs instead of overwriting source data.
- **Formula-driven outputs** — Excel calculations use live formulas instead of hardcoded results where appropriate.
- **Centralized LLM wrapper** — model calls are routed through a single function to make provider changes easier.

## Limitations

- Native Power BI `.pbix` files cannot be generated directly; the agent produces clean exports and a DAX/dashboard guide.
- Native Excel PivotTable objects are not generated; equivalent summary tables are produced instead.
- Excel VBA/macros are not generated.
- The natural-language code execution path uses restricted `exec()` and is intended for personal/local use. A production multi-user deployment would require stronger sandboxing.

## Development Note

This is an **AI-assisted learning project**. Claude was used extensively during development for code generation, debugging, refactoring, and implementation guidance. The project requirements, feature direction, testing, and final integration were reviewed and directed by the author.

The purpose of the project is to learn how LLMs can be integrated into practical data-analysis workflows rather than to present the implementation as entirely hand-written code.

## Future Improvements

- Web interface using Streamlit or Gradio
- Multi-file joins and cross-file analysis
- Natural-language chart generation
- Write support for Word and PowerPoint files
- Stronger sandboxing for generated code

## Author

Sumeet Rawat
