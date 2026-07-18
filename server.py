"""
Docling MCP Server
-------------------
Exposes converter.py's project-conversion pipeline as MCP tools, so Claude
can turn a folder of investing documents (10-Ks, investor decks, financial
statements, earnings-call transcripts...) into Markdown / JSON / table CSVs
without spending tokens reading raw PDFs or PPTX files directly.

Why tables/ doesn't try to label tables as "income statement" / "balance
sheet": the equity-research skills this feeds source their structured
financials from Morningstar exports and read 10-Ks/decks/transcripts
narratively, not via extracted table CSVs -- so semantic table
classification isn't something the actual workflow needs. tables/ still
gets a plain per-table CSV (table_NN.csv + source page number) as a cheap,
useful-if-you-need-it export, but there's no attempt to guess which table
is which -- see converter.py's write_tables() docstring.

Project layout (see converter.py's module docstring for full detail):
    <PROJECTS_DIR>/<project_name>/
        source/       <- put the raw files here yourself
        processed/    <- created by convert_project
        metadata.json <- read this first (read_project_metadata)

PROJECTS_DIR is fixed to a "projects" folder next to this script. Every
tool call resolves project_name relative to that root and refuses anything
that could escape it (no "..", no path separators, no absolute paths) --
see _resolve_project_dir().

Setup
-----
    pip install -r requirements.txt

Add this server to your MCP client's config. For Claude Desktop on
Windows, that's usually %APPDATA%\\Claude\\claude_desktop_config.json
(Settings -> Developer -> Edit Config opens it) -- add an entry like:

    {
      "mcpServers": {
        "docling": {
          "command": "C:\\path\\to\\venv\\Scripts\\python.exe",
          "args": ["C:\\Documents\\PersonalProjects\\Claude\\Docling\\MCP\\server.py"]
        }
      }
    }

Use the full path to your venv's python.exe (not just "python") -- Claude
Desktop launches the config with a minimal PATH and short names often
don't resolve. Restart Claude Desktop after editing the config. Then drop
files into <this folder>/projects/<ProjectName>/source/ and ask Claude to
convert that project.

Note: legacy .xls/.xlsb/.xlsm conversion shells out to Excel via
win32com (see converter.py), so this server needs to run on Windows with
Excel installed and available to automate -- fine for a local desktop
setup like this one, but worth knowing if you ever move it elsewhere.

Long-running conversions
-------------------------
A real conversion batch can take up to ~10 minutes (first-run Docling
model download, OCR-heavy pages, lots of files). convert_project does NOT
block until the job is done -- it starts the job in the background and
returns, because a call that blocked for the full duration would always
eventually exceed some client-side timeout. In practice, this setup's
client gives up on a single tool call after about 4 minutes ("No result
received... after waiting 4 minutes") -- so no single call here should
ever run longer than that, including a long-poll wait.

Both convert_project and check_conversion_status accept an optional
wait_seconds: if > 0, the call itself waits (cheap in-process polling,
not separate tool calls) for up to that many seconds for the job to
finish, capped at MAX_WAIT_SECONDS (comfortably under the ~4-minute
ceiling above) before returning whatever the status is at that point.
This exists specifically to keep token cost down: checking a 10-minute
job with wait_seconds=0 every time means dozens of near-instant round
trips; using wait_seconds=MAX_WAIT_SECONDS means roughly
ceil(total_time / MAX_WAIT_SECONDS) calls instead -- about 2-3 for a
10-minute job.

This in-progress job tracking lives in memory (_active_jobs below) and
only persists for as long as this server process stays alive -- if you
restart Claude Desktop (which restarts this server subprocess) while a
conversion is running, that job is genuinely gone, not just untracked.
Don't restart mid-conversion; long-poll instead. A restart AFTER a job
finishes is fine -- its result is already durably on disk in
metadata.json, which check_conversion_status and read_project_metadata
both fall back to when there's no in-memory job to report.
"""

import asyncio
import json
import time
from datetime import datetime
from pathlib import Path

from mcp.server.fastmcp import FastMCP
from pydantic import BaseModel, ConfigDict, Field

import converter

mcp = FastMCP("docling_mcp")

SCRIPT_DIR = Path(__file__).resolve().parent
PROJECTS_DIR = SCRIPT_DIR / "projects"
PROJECTS_DIR.mkdir(exist_ok=True)

# This setup's MCP client has been observed to give up on a single tool
# call after ~4 minutes (240s). MAX_WAIT_SECONDS keeps every long-poll
# call comfortably under that, so waiting itself never becomes the thing
# that times out. POLL_INTERVAL_SECONDS is how often the wait loop checks
# job status internally while waiting -- an implementation detail, not
# something a caller needs to think about.
MAX_WAIT_SECONDS = 220
POLL_INTERVAL_SECONDS = 3

# In-memory registry of background conversion jobs, keyed by project_name.
# See "Long-running conversions" above for what this does and doesn't survive.
#   {"status": "running"|"success"|"error", "started_at": str,
#    "started_monotonic": float, "finished_at": str|None,
#    "result": dict|None, "error": str|None, "task": asyncio.Task}
_active_jobs: dict[str, dict] = {}


def _resolve_project_dir(project_name: str) -> Path:
    """
    Resolves project_name to a folder inside PROJECTS_DIR, refusing anything
    that could escape that root. Raises ValueError with an actionable
    message on violation -- callers turn that into a JSON {"error": ...}
    response rather than letting it become an unhandled exception.
    """
    if not project_name or not project_name.strip():
        raise ValueError("project_name cannot be empty.")
    if "/" in project_name or "\\" in project_name or ".." in project_name:
        raise ValueError(
            f"project_name must be a plain folder name (e.g. 'Tesla_Project'), "
            f"not a path. Got: {project_name!r}"
        )
    root = PROJECTS_DIR.resolve()
    candidate = (root / project_name).resolve()
    if candidate != root and root not in candidate.parents:
        raise ValueError(f"project_name resolves outside the projects/ root: {project_name!r}")
    return candidate


_WAIT_SECONDS_DESCRIPTION = (
    "If > 0, this call waits (cheaply, in-process -- not separate tool "
    "calls) for up to this many seconds for the job to leave 'running' "
    f"state before returning, capped at {MAX_WAIT_SECONDS}. Use this "
    "instead of firing many quick checks: a conversion batch can take up "
    f"to ~10 minutes, and each separate check has a fixed token cost, so "
    f"waiting {MAX_WAIT_SECONDS}s per call means roughly 2-3 calls total "
    "for a 10-minute job instead of dozens of instant ones. Returns "
    "early the moment the job finishes, so a short conversion with a "
    "long wait_seconds still returns promptly. Default 0 = check once, "
    "instantly, no waiting."
)


class ProjectNameInput(BaseModel):
    """Input model shared by tools that only need to identify a project."""
    model_config = ConfigDict(str_strip_whitespace=True, extra="forbid")

    project_name: str = Field(
        ...,
        description=(
            "Folder name of the project under projects/ (e.g. 'Tesla_Project'). "
            "Must already exist with a source/ subfolder containing the files "
            "to convert. Use list_projects to see what's available."
        ),
        min_length=1,
        max_length=200,
    )


class CheckStatusInput(BaseModel):
    """Input model for check_conversion_status."""
    model_config = ConfigDict(str_strip_whitespace=True, extra="forbid")

    project_name: str = Field(
        ...,
        description=(
            "Folder name of the project under projects/ (e.g. 'Tesla_Project')."
        ),
        min_length=1,
        max_length=200,
    )
    wait_seconds: int = Field(
        default=0,
        description=_WAIT_SECONDS_DESCRIPTION,
        ge=0,
        le=MAX_WAIT_SECONDS,
    )


class ConvertProjectInput(BaseModel):
    """Input model for convert_project."""
    model_config = ConfigDict(str_strip_whitespace=True, extra="forbid")

    project_name: str = Field(
        ...,
        description=(
            "Folder name of the project under projects/ (e.g. 'Tesla_Project'). "
            "Must already contain a source/ subfolder with the files to convert."
        ),
        min_length=1,
        max_length=200,
    )
    chunk_threshold: int = Field(
        default=25,
        description=(
            "Pages/slides/blocks above which a file is split into chunks "
            "before conversion, to stop Docling from silently truncating "
            "large documents. Files at or under this size convert whole, "
            "which is required for JSON/table export -- most 10-Ks and "
            "long decks will exceed this and produce Markdown only."
        ),
        ge=1,
        le=1000,
    )
    force: bool = Field(
        default=False,
        description=(
            "If true, reconverts every file even if it already converted "
            "successfully last run and hasn't changed since. If false "
            "(default), only new or modified files are (re)converted -- "
            "much faster for a project that's already been processed."
        ),
    )
    wait_seconds: int = Field(
        default=0,
        description=_WAIT_SECONDS_DESCRIPTION + (
            " Applies whether this call starts a fresh job or finds one "
            "already running for this project -- either way, it waits on "
            "whatever job ends up active."
        ),
        ge=0,
        le=MAX_WAIT_SECONDS,
    )


async def _wait_for_job(project_name: str, wait_seconds: int) -> dict:
    """
    Polls _active_jobs[project_name] internally -- cheap in-process checks,
    not separate tool calls -- until it leaves "running" state or
    wait_seconds elapses, whichever comes first. Returns the job dict as it
    stood at that point. Shared by convert_project (optional wait right
    after starting) and check_conversion_status (explicit long-poll), so
    either can turn many quick round trips into one longer one.
    """
    deadline = time.monotonic() + wait_seconds
    while True:
        job = _active_jobs[project_name]
        if job["status"] != "running" or time.monotonic() >= deadline:
            return job
        await asyncio.sleep(min(POLL_INTERVAL_SECONDS, max(0.0, deadline - time.monotonic())))


def _job_status_response(job: dict) -> dict:
    """Formats a job dict from _active_jobs into the JSON-able response
    shape shared by convert_project and check_conversion_status."""
    if job["status"] == "running":
        return {
            "status": "running",
            "started_at": job["started_at"],
            "elapsed_seconds": round(time.monotonic() - job["started_monotonic"], 1),
        }
    if job["status"] == "success":
        return {
            "status": "success",
            "started_at": job["started_at"],
            "finished_at": job["finished_at"],
            "result": job["result"],
        }
    return {
        "status": "error",
        "started_at": job["started_at"],
        "finished_at": job["finished_at"],
        "error": job["error"],
    }


@mcp.tool(
    name="list_projects",
    annotations={
        "title": "List Docling Projects",
        "readOnlyHint": True,
        "destructiveHint": False,
        "idempotentHint": True,
        "openWorldHint": False,
    },
)
async def list_projects() -> str:
    """Lists every project folder under projects/, with a quick status for
    each: whether it has a source/ folder, whether it's been converted yet,
    and if so, how many files and when it last ran. Call this first when
    you don't already know the exact project_name, or to check what's
    available before deciding whether to convert or just read.

    Returns:
        str: JSON array of objects, one per project folder:
            {
              "project_name": str,
              "has_source": bool,
              "converted": bool,
              "files_processed": int | null,
              "summary": {"success": int, "partial": int, "failed": int} | null,
              "last_run": str | null,
              "currently_running": bool
            }
    """
    results = []
    for entry in sorted(PROJECTS_DIR.iterdir()):
        if not entry.is_dir():
            continue
        meta = converter.read_metadata(entry)
        job = _active_jobs.get(entry.name)
        results.append(
            {
                "project_name": entry.name,
                "has_source": (entry / "source").is_dir(),
                "converted": meta is not None,
                "files_processed": meta["files_processed"] if meta else None,
                "summary": meta["summary"] if meta else None,
                "last_run": meta["last_run"] if meta else None,
                "currently_running": job is not None and job["status"] == "running",
            }
        )
    return json.dumps(results, indent=2)


@mcp.tool(
    name="read_project_metadata",
    annotations={
        "title": "Read Project Metadata",
        "readOnlyHint": True,
        "destructiveHint": False,
        "idempotentHint": True,
        "openWorldHint": False,
    },
)
async def read_project_metadata(params: ProjectNameInput) -> str:
    """Reads a project's metadata.json without converting anything. Use this
    to check what's already been converted -- file paths, status, errors --
    before deciding whether convert_project needs to run at all.

    Args:
        params (ProjectNameInput): project_name identifying the project folder.

    Returns:
        str: JSON object matching metadata.json's schema, or {"error": str}
        if the project or its metadata.json doesn't exist yet:
            {
              "project": str,
              "source_folder": str,
              "files_processed": int,
              "files": [
                {
                  "name": str, "type": str,
                  "status": "success" | "partial" | "failed",
                  "markdown": str | null, "json": str | null,
                  "tables": [{"file": str, "page": int | null}],
                  "chunked": bool, "source_mtime": float,
                  "converted_at": str, "errors": [str]
                }
              ],
              "skipped_unsupported": [str],
              "summary": {"success": int, "partial": int, "failed": int},
              "chunk_threshold": int,
              "last_run": str
            }
    """
    try:
        project_dir = _resolve_project_dir(params.project_name)
    except ValueError as e:
        return json.dumps({"error": str(e)})

    meta = converter.read_metadata(project_dir)
    if meta is None:
        if project_dir.is_dir():
            msg = (
                f"No metadata.json yet for project {params.project_name!r} -- "
                f"it hasn't been converted yet. Call convert_project first."
            )
        else:
            msg = (
                f"Project {params.project_name!r} doesn't exist under projects/. "
                f"Call list_projects to see what's available."
            )
        return json.dumps({"error": msg})

    return json.dumps(meta, indent=2)


@mcp.tool(
    name="convert_project",
    annotations={
        "title": "Convert Project Documents",
        "readOnlyHint": False,
        "destructiveHint": False,
        "idempotentHint": True,
        "openWorldHint": False,
    },
)
async def convert_project(params: ConvertProjectInput) -> str:
    """Starts converting every supported file in a project's source/ folder
    to Markdown (and, for whole/unchunked documents, JSON + table CSVs).

    By default (wait_seconds=0) this does NOT wait for the conversion to
    finish -- it starts the job in the background and returns almost
    immediately with status "running", "success" (if it happens to finish
    within wait_seconds), or "error". Pass wait_seconds > 0 to have this
    same call wait longer before returning -- see wait_seconds' own
    description for why that's cheaper than polling with many separate
    calls. Either way, call check_conversion_status afterward if this
    returns "running" and you want the eventual result.

    Calling this again on a project already mid-conversion does not start a
    second overlapping job -- it waits on (and reports) the one already
    running. Calling it again on a project that already finished
    (successfully or not) starts a fresh run; by default only new/changed
    files are (re)converted (see the force parameter), so re-running on an
    already-converted project finishes fast.

    Supported input types: PDF, DOCX, PPTX, XLSX, HTML, AsciiDoc, PNG, JPG,
    JPEG, TIFF, BMP, WEBP, plus legacy XLS/XLSB/XLSM (auto-converted to
    XLSX first -- requires Excel installed on this machine).

    Args:
        params (ConvertProjectInput): validated input containing:
            - project_name (str): the project folder under projects/
            - chunk_threshold (int): split threshold in pages/slides/blocks
            - force (bool): reconvert everything vs. only what changed
            - wait_seconds (int): how long this call may wait before
              returning whatever the status is at that point

    Returns:
        str: JSON object, one of:
            {"status": "running", "started_at": str, "elapsed_seconds": float,
             "project_name": str, "newly_started": bool, "note": str}
            {"status": "success", "started_at": str, "finished_at": str,
             "result": <metadata.json contents>, "project_name": str, "newly_started": bool}
            {"status": "error", "started_at": str, "finished_at": str,
             "error": str, "project_name": str, "newly_started": bool}
            {"error": str}   -- invalid project_name, or no source/ folder yet
    """
    try:
        project_dir = _resolve_project_dir(params.project_name)
    except ValueError as e:
        return json.dumps({"error": str(e)})

    if not (project_dir / "source").is_dir():
        return json.dumps(
            {
                "error": (
                    f"No source/ folder found for project {params.project_name!r}. "
                    f"Create {project_dir / 'source'} and add the files to convert, "
                    f"then try again."
                )
            }
        )

    existing = _active_jobs.get(params.project_name)
    newly_started = existing is None or existing["status"] != "running"

    if newly_started:

        def _on_done(task: "asyncio.Task", project_name: str = params.project_name) -> None:
            job = _active_jobs[project_name]
            job["finished_at"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            if task.cancelled():
                job["status"] = "error"
                job["error"] = "Conversion was cancelled (the server likely restarted mid-job)."
                return
            exc = task.exception()
            if exc is not None:
                job["status"] = "error"
                job["error"] = f"{type(exc).__name__}: {exc}"
            else:
                job["status"] = "success"
                job["result"] = task.result()

        # convert_folder is blocking (Docling conversion + file I/O). Running
        # it as a Task wrapping asyncio.to_thread is what lets this function
        # return without waiting for it -- a client-side timeout on any one
        # call no longer has any bearing on whether the conversion completes.
        task = asyncio.create_task(
            asyncio.to_thread(
                converter.convert_folder, project_dir, params.chunk_threshold, params.force
            )
        )
        _active_jobs[params.project_name] = {
            "status": "running",
            "started_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "started_monotonic": time.monotonic(),
            "finished_at": None,
            "result": None,
            "error": None,
            "task": task,  # keep a reference -- asyncio only holds tasks weakly otherwise
        }
        task.add_done_callback(_on_done)

    job = await _wait_for_job(params.project_name, params.wait_seconds)
    resp = _job_status_response(job)
    resp["project_name"] = params.project_name
    resp["newly_started"] = newly_started
    if resp["status"] == "running":
        resp["note"] = (
            "Still running. Call check_conversion_status (with wait_seconds "
            "if you want to keep long-polling) rather than calling "
            "convert_project again."
        )
    return json.dumps(resp, indent=2)


@mcp.tool(
    name="check_conversion_status",
    annotations={
        "title": "Check Conversion Status",
        "readOnlyHint": True,
        "destructiveHint": False,
        "idempotentHint": True,
        "openWorldHint": False,
    },
)
async def check_conversion_status(params: CheckStatusInput) -> str:
    """Checks whether a conversion started by convert_project is still
    running, finished successfully, or errored. With the default
    wait_seconds=0, returns instantly. With wait_seconds > 0, waits up to
    that long (cheaply, in-process) for the job to finish before
    returning -- prefer this over calling with wait_seconds=0 many times
    in a row, since each separate call has a fixed token cost regardless
    of how little changed between checks.

    Args:
        params (CheckStatusInput): project_name to check, and optional
            wait_seconds (see its own description for why/how to use it).

    Returns:
        str: JSON object, one of:
            {"status": "running", "started_at": str, "elapsed_seconds": float}
            {"status": "success", "started_at": str, "finished_at": str,
             "result": <metadata.json contents, see read_project_metadata's docstring>}
            {"status": "error", "started_at": str, "finished_at": str, "error": str}
            {"status": "no_active_job", "on_disk_metadata": <metadata.json contents | null>}
                -- no conversion has been started in THIS server session
                (e.g. it finished in a previous session, or never ran, or the
                server restarted mid-job -- see module docstring). Falls back
                to whatever's already on disk, if anything. wait_seconds has
                no effect here -- there's no in-progress job to wait on.
        Or {"error": str} if project_name is invalid.
    """
    try:
        project_dir = _resolve_project_dir(params.project_name)
    except ValueError as e:
        return json.dumps({"error": str(e)})

    if params.project_name not in _active_jobs:
        return json.dumps(
            {
                "status": "no_active_job",
                "on_disk_metadata": converter.read_metadata(project_dir),
            },
            indent=2,
        )

    job = await _wait_for_job(params.project_name, params.wait_seconds)
    return json.dumps(_job_status_response(job), indent=2)


if __name__ == "__main__":
    mcp.run()
