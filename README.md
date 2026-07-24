# Data Analyst AI Agent (Free Gemini Version)

Ek AI agent jo **kisi bhi file** — CSV, Excel, PDF, Word, PowerPoint — ko padhta, samajhta, aur analyze karta hai natural language se. **Google Gemini ka free tier use karta hai — koi credit card ya payment nahi chahiye.**

## Features
- **Universal File Reader**: `.csv`, `.xlsx`, `.pdf`, `.docx`, `.pptx`
- **Auto Cleaning**: file load hote hi generic cleaning
- **Custom Cleaning**: `clean: <instruction>` se apni marzi ki cleaning (session ke liye persist)
- **Save**: `save` command se cleaned data CSV + Excel dono mein export
- **SQL Mode**: `sql: <question>` se query generate + run
- **Persistent SQL DB**: `sql-export` se real `.db` file — SQL Workbench/DB Browser mein khul sakti hai, CREATE TABLE aur data dono visible
- **Auto EDA Report**: `report` command
- **Excel Dashboard**: `dashboard` — charts + dynamic formulas
- **Power BI Export**: `powerbi` — clean, import-ready data
- **Power BI Guide**: `powerbi-guide` — DAX measures + dashboard-building steps

## Setup (Free — No Card Needed)

1. **https://aistudio.google.com/apikey** pe free key lo (Google login, "Create API Key")
2. Libraries install karo:
```bash
pip install google-genai pandas openpyxl python-docx python-pptx pypdf
```
3. Key set karo:

**Windows (PowerShell):**
```powershell
$env:GEMINI_API_KEY="your-key-yaha-daalo"
```
**Mac/Linux:**
```bash
export GEMINI_API_KEY="your-key-yaha-daalo"
```

## Run

```bash
python agent.py your_data.csv       # ya .xlsx, .pdf, .docx, .pptx
```

### Tabular files (csv/xlsx) commands:
```
You: report
You: clean: negative sales values hatao
You: clean: Region blank ho to 'Unknown' likho
You: save                    # cleaned data ko CSV + Excel mein save karega
You: sql: top 5 products by revenue
You: sql-export              # real .db file banata hai — SQL Workbench mein khol sakte ho
You: dashboard                # Excel: charts + dynamic formulas
You: powerbi                  # Power BI-ready data
You: powerbi-guide            # DAX measures + dashboard guide
You: exit
```

**Typical workflow:** messy file do → `clean: ...` (jitni baar chahiye) → `save` (final cleaned file mil jayegi) → `sql-export` (agar DB chahiye external tools ke liye) → `dashboard`/`powerbi` (reporting ke liye).

### PDF/Word/PPT ke liye:
```
You: is document ka summary do
You: is report mein revenue ke baare mein kya likha hai?
You: exit
```

## Free tier limits
Gemini free tier mein daily request limit hai (personal use ke liye kaafi generous). "Quota exceeded" aaye to thodi der wait karo.

## Power BI note
Agent `.pbix` file khud nahi bana sakta (koi public API nahi hai). Deta hai: clean data, DAX measures, step-by-step guide — dashboard Power BI Desktop mein khud banana hoga.

## Portfolio tip
Resume bullet: "Built a multi-format AI data analyst agent (Python + Gemini API) — automated cleaning, NL-to-SQL querying, persistent SQLite export, Excel dashboards, and Power BI-ready exports with DAX generation."

## Security note
`exec()` restricted builtins ke saath use hota hai — untrusted users ke liye production mein aur sandbox karna padega.
