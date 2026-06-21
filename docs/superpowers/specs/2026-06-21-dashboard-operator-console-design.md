# Dashboard Operator Console Design

## Goal

Create a local browser dashboard that lets the user run the existing content scripts from buttons instead of manually opening PowerShell in each script folder.

The dashboard will control two workflows:

- Single Hadith Content: `SINGLE HADITH CONTENT/reconcile_hadith_outputs.py`
- Quran Content: `QURAN CONTENT/convert_quran_content.py`

## User Experience

The first screen is an operator console, not a landing page. A left sidebar lets the user switch between `Single Hadith Content` and `Quran Content`. The main panel shows the selected workflow, its current status, expected input and output folders, controls, and recent run output.

Each workflow has a primary `Run` button. The Hadith workflow also has a `Dry run` toggle because the script already supports `--dry-run`. Quran conversion runs normally because its CLI does not currently expose a dry-run mode.

The console shows live logs while a script is running, then shows success or failure when the process exits. After completion, the UI lists the most relevant generated files and report folders.

## Architecture

Add a small Python dashboard server at the workspace root. It will serve the frontend assets and expose local HTTP endpoints for starting jobs, reading job status, and streaming logs.

The server runs each existing script as a subprocess with its working directory set to the script folder. This preserves the scripts' current relative defaults:

- Hadith script runs from `SINGLE HADITH CONTENT`
- Quran script runs from `QURAN CONTENT`

The dashboard server should not move or rewrite the existing scripts. It is a thin control layer over the current tested tools.

## Components

### Dashboard Server

Responsibilities:

- Serve the dashboard HTML, CSS, and JavaScript.
- Start a selected workflow as a subprocess.
- Track whether a workflow is idle, running, succeeded, or failed.
- Capture stdout and stderr into an in-memory log buffer.
- Prevent starting a workflow while that same workflow is already running.
- Return recent output artifact paths for the selected workflow.

### Frontend

Responsibilities:

- Render the operator console layout.
- Switch between Hadith and Quran workflows.
- Show workflow-specific controls.
- Disable run controls while the selected workflow is running.
- Poll the server for status and logs.
- Show clear success and failure states.

### Workflow Registry

The server keeps one registry of supported workflows. Each workflow entry defines:

- display name
- script path
- working directory
- default arguments
- optional arguments controlled by the UI
- important input folders
- important output/report paths

This keeps command construction explicit and testable.

## Data Flow

1. User opens the dashboard in a browser.
2. Browser requests workflow metadata from the Python dashboard server.
3. User selects Hadith or Quran and clicks `Run`.
4. Browser sends a start request for that workflow.
5. Server builds the command and starts the script subprocess in the correct working directory.
6. Server captures stdout and stderr as log lines.
7. Browser polls for job status and log updates.
8. When the process exits, the server stores the final exit code and output summary.
9. Browser shows success or failure and links to relevant generated files or folders.

## Error Handling

If a script exits with a nonzero status, the dashboard marks the run as failed and keeps the final log output visible.

If the user tries to run a workflow that is already running, the server rejects the request and the UI keeps the existing run visible.

If Python or a script file cannot be found, the server returns a clear startup error instead of silently failing.

If the browser loses connection while a script continues running, refreshing the page shows the latest known status and logs.

## Testing

Add automated tests for the dashboard server logic before implementation:

- workflow registry contains the Hadith and Quran workflows
- Hadith command includes `--dry-run` only when requested
- subprocesses are started with the correct working directory
- a second start request is rejected while a workflow is running
- successful and failed exits update status correctly

Manual verification:

- Start the dashboard server.
- Open the local URL.
- Run Hadith in dry-run mode and confirm logs appear.
- Run Quran conversion and confirm output summary appears.

## Out of Scope

- Editing workbook files from the dashboard.
- Choosing arbitrary custom folders in the first version.
- Running both workflows in parallel from one combined button.
- Replacing the existing script CLIs.
