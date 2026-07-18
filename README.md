# Docling MCP Server

Turns a folder of investing documents (10-Ks, investor decks, financial
statements, earnings-call transcripts, etc.) into Markdown / JSON / table
CSVs that Claude can read cheaply, instead of opening raw PDFs/PPTX files
directly every time.

Built from your original `docling_batch_convert_interactive.py`, split into:

- **`converter.py`** — the conversion engine (your original Docling +
  chunking logic, unchanged, plus a new `convert_folder()` entry point and
  a generic table-CSV exporter). No prompts, no CLI — pure library code.
- **`server.py`** — the MCP layer. Exposes `convert_folder()` as three MCP
  tools Claude can call.

## Folder layout

Every "project" lives in its own folder under `projects/` (created
automatically next to `server.py`):

```
MCP/
├── converter.py
├── server.py
├── requirements.txt
└── projects/
    └── Tesla_Project/
        ├── source/                 <- you put raw files here
        │   ├── Annual Report.pdf
        │   ├── Investor Presentation.pptx
        │   └── Financial Statements.xlsx
        ├── processed/               <- created automatically
        │   ├── markdown/
        │   │   ├── Annual Report.md
        │   │   └── Investor Presentation.md
        │   ├── json/
        │   │   └── Financial Statements.json
        │   └── tables/
        │       └── Financial Statements__table_01.csv
        └── metadata.json            <- Claude reads this first
```

`source/` is never modified — files are read in place, never moved or
archived. Re-running conversion on a project only (re)converts files that
are new or have changed since the last run; everything else is reused from
`metadata.json`.

**Docling can't tell which table is which statement on its own. 
So `tables/` holds a plain, generic CSV per table
(with a page-number reference in `metadata.json`) —
useful if you ever need it, without a classification layer your workflow
doesn't currently use.

**Why some files only get Markdown, no JSON/tables:** a file above
`chunk_threshold` (default 25 pages/slides/blocks) gets split into chunks
before conversion, and a chunked file has no single combined document
object to export as JSON or tables from — only the stitched Markdown. Most
10-Ks and long decks will hit this. That's fine for your actual workflow
(both skills read Markdown/PDF narratively), just worth knowing.

## Setup

1. **Install dependencies** (Windows, with Excel installed for legacy
   `.xls`/`.xlsb`/`.xlsm` support):

   ```
   cd C:\Documents\PersonalProjects\Claude\Docling\MCP
   python -m venv venv
   venv\Scripts\activate
   pip install -r requirements.txt
   ```

2. **Add the server to your MCP client's config.** For Claude Desktop on
   Windows, that's usually:

   ```
   %APPDATA%\Claude\claude_desktop_config.json
   ```

   Open it via Claude Desktop's **Settings → Developer → Edit Config**
   (more reliable than guessing the path yourself — some Windows installs
   store it at a different, virtualized location, and "Edit Config" opens
   the right one). Add:

   ```json
   {
     "mcpServers": {
       "docling": {
         "command": "C:\\Documents\\PersonalProjects\\Claude\\Docling\\MCP\\venv\\Scripts\\python.exe",
         "args": ["C:\\Documents\\PersonalProjects\\Claude\\Docling\\MCP\\server.py"]
       }
     }
   }
   ```

   Use the **full path** to `venv\Scripts\python.exe`, not just `python` —
   Claude Desktop launches the config with a minimal PATH and short names
   often don't resolve. If `claude_desktop_config.json` already has other
   content (other servers, preferences), merge `"docling"` into the
   existing `"mcpServers"` object rather than replacing the file.

3. **Restart Claude Desktop** fully (quit, not just close the window).

4. Check **Settings → Developer** to confirm the `docling` server shows as
   connected. If it doesn't, the same screen has logs to check.

## Using it

Drop files into `projects/<ProjectName>/source/` yourself (create the
folder if it's a new project), then just ask Claude to convert it — it has
four tools:

- **`list_projects`** — see what projects exist, their status, and whether
  a conversion is currently running.
- **`convert_project`** — starts converting a project's `source/` into
  `processed/` + `metadata.json`. Takes an optional `wait_seconds`
  (0-220, default 0): with 0, it returns almost immediately once the job
  has *started*, without waiting for it to finish; with a higher value,
  it waits up to that long *in that same call* before returning, and
  returns early the moment the job actually finishes. Use a non-zero
  `wait_seconds` to avoid burning tokens on repeated status checks — see
  "Avoiding excess polling" below.
- **`check_conversion_status`** — checks a project's in-progress (or just
  finished) conversion. Same `wait_seconds` parameter, same reasoning.
- **`read_project_metadata`** — reads the last *completed* run's
  `metadata.json` without touching anything currently in progress.

### Avoiding excess polling

A real conversion batch can take up to ~10 minutes. Checking on it with
`wait_seconds=0` every time means many small tool calls, each with its
own fixed token cost, even though nothing's changed between most of them.
Instead, pass a high `wait_seconds` (capped at 220, just under the ~4
minute ceiling this MCP client appears to enforce on any single call) so
one call absorbs the wait:

```
convert_project(project_name="Tesla_Project", wait_seconds=220)
```

If it's still running after 220s, that call returns `{"status": "running", ...}`
and you call `check_conversion_status(project_name="Tesla_Project", wait_seconds=220)`
again — for a ~10-minute job, that's roughly 2-3 calls total instead of
dozens of instant ones. If it finishes sooner, the call returns as soon as
it does, not after the full 220s.

Example: "Convert the Tesla_Project folder (wait up to 220s), then read
the 10-K and give me the business model, financials, valuation, and
risks" — Claude will call `convert_project` with a long `wait_seconds`,
long-poll `check_conversion_status` only if it's still running after
that, then read the resulting Markdown directly instead of the raw PDF.

**Do not restart Claude Desktop while a conversion is running.** Progress
tracking lives in the server process's memory — restarting the app
restarts that process and genuinely kills whatever was mid-conversion, not
just the connection to it. If a job seems stuck, check Task Manager first
(see below) rather than restarting; a restart is only safe once
`check_conversion_status` shows `"success"` or `"error"`, or you're
willing to lose that in-flight run and start over (already-converted files
are still skipped on the next attempt, per `metadata.json`).

## Known limitations

- Legacy Excel conversion (`.xls`/`.xlsb`/`.xlsm`) requires Windows with
  Excel installed and automatable (no other automation locking Excel at
  the same time). A hidden Excel instance can still pop a modal dialog
  that blocks forever with nothing visible to click — if a job seems
  stuck and `source/` has a legacy Excel file, check Task Manager for an
  `EXCEL.EXE` process with near-zero CPU and end it if found.
- `project_name` is restricted to a plain folder name under `projects/` —
  no path separators or `..`, so a tool call can't reach outside that root.
- `wait_seconds` is capped at 220s per call based on one observed timeout
  on this setup (~4 minutes) — if your client turns out to tolerate more
  or less, adjust `MAX_WAIT_SECONDS` near the top of `server.py`.
- Conversion progress tracking only lives in the server process's memory
  (see above) — it doesn't survive a server restart, though completed
  results in `metadata.json` do.
