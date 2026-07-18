# Docling MCP Server

## AI Optimised Document Conversion for Financial Research Libraries

This MCP server converts complete folders of investment and financial documents into structured formats that Claude and other AI assistants can efficiently navigate.

It transforms:

- Annual reports
- 10-K / 20-F filings
- Investor presentations
- Earnings call transcripts
- Financial statements
- Excel financial models
- Research documents

into:

- Markdown files for AI reasoning
- JSON files for structured extraction
- CSV files for financial tables
- Metadata indexes for fast AI navigation


---

# Why this exists

The standard Docling MCP workflow works well for individual documents.

However, when Claude is connected directly to a general Docling MCP server and asked to process:

- Large annual reports
- Multiple financial documents
- Entire research folders
- Complex investor presentations
- Large Excel workbooks

several issues appear:

- Large files consume excessive context and tokens.
- Multiple files can overload the MCP connection.
- Claude repeatedly processes the same source documents.
- Financial tables are difficult for AI models to locate efficiently.
- Older Excel formats require special handling.
- The AI must search through entire documents before finding relevant information.

This MCP server was created to solve those problems.

Instead of Claude repeatedly opening raw files, this server creates an AI readable research database.

The workflow becomes:

```
Raw documents
        |
        v
Docling MCP conversion
        |
        v
Structured research database
        |
        v
Claude navigates metadata
        |
        v
Claude reads only relevant files
```


---

# Key Features

## 1. Convert an entire folder automatically

Instead of processing one document at a time, point Claude to a project folder.

Example:

```
Tesla_Project/

├── Annual Report.pdf
├── Investor Presentation.pptx
├── Financial Statements.xlsx
├── Earnings Call Transcript.pdf
└── Research Notes.docx
```

The server automatically processes the folder.


---

## 2. Automatic output selection

The server chooses the best output format depending on the document type.

Example:

| Input File | Output |
|---|---|
| Annual Report PDF | Markdown + extracted tables |
| Investor Presentation PPTX | Markdown |
| Financial Statements XLSX | JSON + CSV tables |
| Legacy Excel XLS/XLSB/XLSM | Converted automatically |
| Text documents | Markdown |


---

## 3. Metadata driven AI navigation

Every project creates:

```
metadata.json
```

This acts as a map for Claude.

Instead of:

```
Claude
 |
 opens every PDF
 |
 searches thousands of pages
```

The workflow becomes:

```
Claude
 |
 reads metadata.json
 |
 identifies relevant document
 |
 opens required Markdown/JSON/table
 |
 performs analysis
```

This reduces:

- Context usage
- Token consumption
- Processing time


---

## 4. Financial report optimisation

Financial reports contain important structures:

- Income statements
- Balance sheets
- Cash flow statements
- Segment reporting
- Revenue breakdowns
- Debt schedules
- Financial tables


This server preserves these structures in formats designed for AI analysis.


---

# Project Structure

Every research project is stored separately.


```
MCP/

├── converter.py
├── server.py
├── requirements.txt
│
└── projects/

    └── Tesla_Project/

        ├── source/

        │   ├── Annual Report.pdf
        │   ├── Investor Presentation.pptx
        │   └── Financial Statements.xlsx


        └── processed/

            ├── markdown/

            │   ├── Annual Report.md
            │   └── Investor Presentation.md


            ├── json/

            │   └── Financial Statements.json


            ├── tables/

            │   └── Financial Statements__table_01.csv


            └── metadata.json
```


The `source` folder is never modified.

Files remain in their original location.

The server only creates processed outputs.


---

# Incremental Processing

The server tracks converted documents using:

```
metadata.json
```

When conversion runs again:

- New files are converted.
- Modified files are reconverted.
- Existing unchanged files are reused.

Large research libraries do not need full reconversion.


---

# Installation

## Requirements

Recommended:

- Windows
- Python 3.10+
- Excel installed (required for legacy Excel conversion)


---

## Create virtual environment

Navigate to the MCP folder:


```
cd C:\Documents\PersonalProjects\Claude\Docling\MCP
```


Create environment:


```
python -m venv venv
```


Activate:


```
venv\Scripts\activate
```


Install packages:


```
pip install -r requirements.txt
```


---

# Claude Desktop Configuration

Open:

```
Claude Desktop
→ Settings
→ Developer
→ Edit Config
```


Add:

```json
{
  "mcpServers": {
    "docling": {
      "command": "C:\\Documents\\PersonalProjects\\Claude\\Docling\\MCP\\venv\\Scripts\\python.exe",
      "args": [
        "C:\\Documents\\PersonalProjects\\Claude\\Docling\\MCP\\server.py"
      ]
    }
  }
}
```


Important:

Use the full path to:

```
venv\Scripts\python.exe
```

Do not use:

```
python
```

Claude Desktop runs with a limited system PATH.


Restart Claude Desktop completely.


Check:

```
Settings
→ Developer
```

The Docling server should appear as connected.


---

# Recommended Folder Naming

Use clear project names.

Recommended:

```
Tesla_2025
Microsoft_Annual_Report_2024
Nvidia_Investment_Research
```

Avoid:

```
New Folder
Tesla stuff
Documents
Final version
```


The project name becomes the identifier Claude uses when navigating your research library.


---

# Using the MCP Server

The server provides four MCP tools.


## list_projects

Shows:

- Available projects
- Conversion status
- Existing metadata


Example:

```
List available projects
```


---

## convert_project

Converts:

```
projects/<ProjectName>/source/
```

into:

```
projects/<ProjectName>/processed/
```


Example:

```
Convert Tesla_Project
```


Optional:

```
wait_seconds
```

controls how long Claude waits before returning.


Recommended:

```
wait_seconds=220
```


---

## check_conversion_status

Checks:

- Current progress
- Completion status
- Errors


---

## read_project_metadata

Reads:

```
metadata.json
```

without modifying files.


This allows Claude to understand the project structure before opening documents.


---

# Avoiding Excess Token Usage

Large conversions can take several minutes.

Avoid repeatedly asking Claude:

```
Is it finished?
Is it finished?
Is it finished?
```

Each MCP call consumes tokens.


Instead:

```
Convert Tesla_Project and wait up to 220 seconds.
```

If the conversion is still running:

```
Check status and wait another 220 seconds.
```


This reduces unnecessary MCP calls.


---

# Example Workflow

Example prompt:

```
Convert the Tesla_Project folder.

After conversion:
1. Read the annual report.
2. Summarise the business model.
3. Extract financial trends.
4. Analyse valuation risks.
```


Claude will:

1. Start conversion.
2. Wait for completion.
3. Read metadata.
4. Locate relevant documents.
5. Analyse processed files.


---

# Known Limitations


## Legacy Excel files

Support for:

```
.xls
.xlsb
.xlsm
```

requires:

- Windows
- Microsoft Excel installed


Excel must not already be locked by another process.


If conversion appears frozen:

Open Task Manager.

Check:

```
EXCEL.EXE
```

Terminate hidden Excel processes if required.


---

## Large files

Very large documents may be split into chunks.

Chunked documents produce:

- Markdown output

but may not produce:

- Combined JSON
- Combined table exports


This is expected because the original document object no longer exists as a single structure after chunking.


---

## Project names

Project names must be simple folder names.

Allowed:

```
Tesla_2025
```

Not allowed:

```
C:\Research\Tesla
../Tesla
```


This prevents accidental access outside the project directory.


---

## Conversion status

Active conversion progress exists only while the MCP server is running.

Restarting Claude Desktop during conversion will stop the active job.

Completed conversions remain safe because outputs are stored in:

```
metadata.json
```


---

# License

MIT License

Free to use, modify, and distribute.
