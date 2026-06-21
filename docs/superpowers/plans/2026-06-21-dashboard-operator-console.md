# Dashboard Operator Console Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Build and run a local browser dashboard that starts the existing Hadith and Quran scripts from buttons.

**Architecture:** Add a small standard-library Python HTTP server at the workspace root. Keep workflow command construction and process tracking in focused Python classes, and serve static HTML/CSS/JavaScript from a `dashboard/` folder. The frontend polls JSON endpoints for workflow metadata, run status, logs, and output paths.

**Tech Stack:** Python standard library (`http.server`, `subprocess`, `threading`, `json`, `unittest`), browser HTML/CSS/JavaScript, existing Python scripts and `openpyxl`.

---

## File Structure

- Create `dashboard_server.py`: local HTTP server, workflow registry, job manager, JSON API, static file serving.
- Create `dashboard/index.html`: operator console markup.
- Create `dashboard/styles.css`: responsive operator-console visual design.
- Create `dashboard/app.js`: frontend state, run buttons, polling, log rendering.
- Create `tests/test_dashboard_server.py`: server unit tests that avoid running real workbook scripts.

## Task 1: Workflow Registry and Command Construction

**Files:**
- Create: `tests/test_dashboard_server.py`
- Create: `dashboard_server.py`

- [ ] **Step 1: Write failing workflow tests**

Create `tests/test_dashboard_server.py` with:

```python
import sys
import unittest
from pathlib import Path

import dashboard_server as dashboard


class WorkflowRegistryTest(unittest.TestCase):
    def test_registry_contains_hadith_and_quran_workflows(self):
        registry = dashboard.build_workflow_registry(Path.cwd())

        self.assertEqual(set(registry), {"hadith", "quran"})
        self.assertEqual(registry["hadith"].name, "Single Hadith Content")
        self.assertEqual(registry["quran"].name, "Quran Content")
        self.assertEqual(registry["hadith"].working_dir.name, "SINGLE HADITH CONTENT")
        self.assertEqual(registry["quran"].working_dir.name, "QURAN CONTENT")

    def test_hadith_command_adds_dry_run_only_when_requested(self):
        registry = dashboard.build_workflow_registry(Path.cwd())

        normal = registry["hadith"].command({"dry_run": False})
        dry_run = registry["hadith"].command({"dry_run": True})

        self.assertEqual(normal[:2], [sys.executable, "reconcile_hadith_outputs.py"])
        self.assertNotIn("--dry-run", normal)
        self.assertIn("--dry-run", dry_run)

    def test_quran_command_has_no_dry_run_flag(self):
        registry = dashboard.build_workflow_registry(Path.cwd())

        command = registry["quran"].command({"dry_run": True})

        self.assertEqual(command, [sys.executable, "convert_quran_content.py"])
```

- [ ] **Step 2: Run tests to verify failure**

Run: `python -m unittest tests.test_dashboard_server -v`

Expected: import fails because `dashboard_server.py` does not exist.

- [ ] **Step 3: Implement registry and command construction**

Create `dashboard_server.py` with:

```python
from __future__ import annotations

import json
import mimetypes
import subprocess
import sys
import threading
import time
from dataclasses import dataclass, field
from http import HTTPStatus
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from typing import Callable
from urllib.parse import urlparse


ROOT_DIR = Path(__file__).resolve().parent
DASHBOARD_DIR = ROOT_DIR / "dashboard"


@dataclass(frozen=True)
class Workflow:
    id: str
    name: str
    description: str
    script: str
    working_dir: Path
    input_paths: tuple[Path, ...]
    output_paths: tuple[Path, ...]
    supports_dry_run: bool = False

    def command(self, options: dict | None = None) -> list[str]:
        options = options or {}
        command = [sys.executable, self.script]
        if self.supports_dry_run and options.get("dry_run"):
            command.append("--dry-run")
        return command

    def to_dict(self) -> dict:
        return {
            "id": self.id,
            "name": self.name,
            "description": self.description,
            "supports_dry_run": self.supports_dry_run,
            "working_dir": str(self.working_dir),
            "input_paths": [str(path) for path in self.input_paths],
            "output_paths": [str(path) for path in self.output_paths],
        }


def build_workflow_registry(root_dir: Path = ROOT_DIR) -> dict[str, Workflow]:
    hadith_dir = root_dir / "SINGLE HADITH CONTENT"
    quran_dir = root_dir / "QURAN CONTENT"
    return {
        "hadith": Workflow(
            id="hadith",
            name="Single Hadith Content",
            description="Reconcile hadith output workbooks, write reports, and update final content workbooks.",
            script="reconcile_hadith_outputs.py",
            working_dir=hadith_dir,
            input_paths=(hadith_dir / "INPUT", hadith_dir / "OUTPUT", hadith_dir / "BOOKS.xlsx"),
            output_paths=(
                hadith_dir / "reconciled_output",
                hadith_dir / "missing_input",
                hadith_dir / "BOOK WISE FINAL",
                hadith_dir / "BOOK WISE FINAL UPDATED CONTENT",
            ),
            supports_dry_run=True,
        ),
        "quran": Workflow(
            id="quran",
            name="Quran Content",
            description="Convert Quran content workbooks to id, language_id, and contents columns.",
            script="convert_quran_content.py",
            working_dir=quran_dir,
            input_paths=(quran_dir / "INPUT",),
            output_paths=(quran_dir / "OUTPUT", quran_dir / "OUTPUT" / "conversion_summary.csv"),
            supports_dry_run=False,
        ),
    }
```

- [ ] **Step 4: Run tests to verify pass**

Run: `python -m unittest tests.test_dashboard_server -v`

Expected: 3 tests pass.

## Task 2: Job Manager

**Files:**
- Modify: `tests/test_dashboard_server.py`
- Modify: `dashboard_server.py`

- [ ] **Step 1: Add failing job manager tests**

Append to `tests/test_dashboard_server.py`:

```python

class FakeProcess:
    def __init__(self, returncode=0, lines=None):
        self.returncode = returncode
        self.stdout = iter(lines or ["line one\n", "line two\n"])

    def wait(self):
        return self.returncode


class JobManagerTest(unittest.TestCase):
    def test_start_uses_workflow_command_and_working_directory(self):
        calls = []

        def fake_popen(command, **kwargs):
            calls.append((command, kwargs))
            return FakeProcess(returncode=0, lines=["ok\n"])

        registry = dashboard.build_workflow_registry(Path.cwd())
        manager = dashboard.JobManager(registry, popen=fake_popen)

        status = manager.start("hadith", {"dry_run": True})
        manager.wait_for("hadith", timeout=2)

        self.assertEqual(status["state"], "running")
        self.assertEqual(calls[0][0], registry["hadith"].command({"dry_run": True}))
        self.assertEqual(calls[0][1]["cwd"], registry["hadith"].working_dir)
        self.assertEqual(manager.status("hadith")["state"], "succeeded")
        self.assertIn("ok", "\n".join(manager.status("hadith")["logs"]))

    def test_rejects_second_start_while_workflow_is_running(self):
        release = threading.Event()

        class BlockingProcess(FakeProcess):
            def wait(self):
                release.wait(2)
                return 0

        def fake_popen(command, **kwargs):
            return BlockingProcess(lines=["started\n"])

        registry = dashboard.build_workflow_registry(Path.cwd())
        manager = dashboard.JobManager(registry, popen=fake_popen)

        first = manager.start("quran", {})
        second = manager.start("quran", {})
        release.set()
        manager.wait_for("quran", timeout=2)

        self.assertEqual(first["state"], "running")
        self.assertEqual(second["state"], "running")
        self.assertIn("already running", second["message"])

    def test_failed_exit_updates_status(self):
        def fake_popen(command, **kwargs):
            return FakeProcess(returncode=7, lines=["bad\n"])

        registry = dashboard.build_workflow_registry(Path.cwd())
        manager = dashboard.JobManager(registry, popen=fake_popen)

        manager.start("quran", {})
        manager.wait_for("quran", timeout=2)

        status = manager.status("quran")
        self.assertEqual(status["state"], "failed")
        self.assertEqual(status["exit_code"], 7)
```

Also add `import threading` near the top of the test file.

- [ ] **Step 2: Run tests to verify failure**

Run: `python -m unittest tests.test_dashboard_server -v`

Expected: tests fail because `JobManager` is not defined.

- [ ] **Step 3: Implement job manager**

Append to `dashboard_server.py`:

```python
@dataclass
class JobState:
    workflow: Workflow
    state: str = "idle"
    logs: list[str] = field(default_factory=list)
    exit_code: int | None = None
    started_at: float | None = None
    finished_at: float | None = None
    thread: threading.Thread | None = None


class JobManager:
    def __init__(
        self,
        workflows: dict[str, Workflow],
        popen: Callable[..., subprocess.Popen] = subprocess.Popen,
    ):
        self.workflows = workflows
        self.popen = popen
        self._lock = threading.Lock()
        self._jobs = {workflow_id: JobState(workflow) for workflow_id, workflow in workflows.items()}

    def start(self, workflow_id: str, options: dict | None = None) -> dict:
        if workflow_id not in self.workflows:
            return {"state": "failed", "message": f"Unknown workflow: {workflow_id}"}
        options = options or {}
        with self._lock:
            job = self._jobs[workflow_id]
            if job.state == "running":
                return {**self._status_locked(job), "message": f"{job.workflow.name} is already running."}
            job.state = "running"
            job.logs = []
            job.exit_code = None
            job.started_at = time.time()
            job.finished_at = None
            thread = threading.Thread(target=self._run_job, args=(job, options), daemon=True)
            job.thread = thread
            thread.start()
            return self._status_locked(job)

    def status(self, workflow_id: str) -> dict:
        if workflow_id not in self._jobs:
            return {"state": "failed", "message": f"Unknown workflow: {workflow_id}"}
        with self._lock:
            return self._status_locked(self._jobs[workflow_id])

    def wait_for(self, workflow_id: str, timeout: float | None = None) -> None:
        thread = self._jobs[workflow_id].thread
        if thread is not None:
            thread.join(timeout)

    def _run_job(self, job: JobState, options: dict) -> None:
        workflow = job.workflow
        try:
            process = self.popen(
                workflow.command(options),
                cwd=workflow.working_dir,
                stdout=subprocess.PIPE,
                stderr=subprocess.STDOUT,
                text=True,
                encoding="utf-8",
                errors="replace",
                bufsize=1,
            )
            if process.stdout is not None:
                for line in process.stdout:
                    self._append_log(job, line.rstrip())
            exit_code = process.wait()
            with self._lock:
                job.exit_code = exit_code
                job.state = "succeeded" if exit_code == 0 else "failed"
                job.finished_at = time.time()
        except Exception as exc:
            with self._lock:
                job.logs.append(f"Failed to start: {exc}")
                job.exit_code = None
                job.state = "failed"
                job.finished_at = time.time()

    def _append_log(self, job: JobState, line: str) -> None:
        with self._lock:
            job.logs.append(line)
            if len(job.logs) > 500:
                job.logs = job.logs[-500:]

    def _status_locked(self, job: JobState) -> dict:
        return {
            "id": job.workflow.id,
            "name": job.workflow.name,
            "state": job.state,
            "logs": list(job.logs),
            "exit_code": job.exit_code,
            "started_at": job.started_at,
            "finished_at": job.finished_at,
            "outputs": self._output_status(job.workflow),
        }

    def _output_status(self, workflow: Workflow) -> list[dict]:
        outputs = []
        for path in workflow.output_paths:
            outputs.append(
                {
                    "path": str(path),
                    "exists": path.exists(),
                    "is_dir": path.is_dir(),
                }
            )
        return outputs
```

- [ ] **Step 4: Run tests to verify pass**

Run: `python -m unittest tests.test_dashboard_server -v`

Expected: 6 tests pass.

## Task 3: HTTP API and Static Server

**Files:**
- Modify: `tests/test_dashboard_server.py`
- Modify: `dashboard_server.py`

- [ ] **Step 1: Add failing handler factory tests**

Append to `tests/test_dashboard_server.py`:

```python

class ApiResponseTest(unittest.TestCase):
    def test_api_workflows_returns_serializable_workflow_list(self):
        registry = dashboard.build_workflow_registry(Path.cwd())
        manager = dashboard.JobManager(registry)

        payload = dashboard.api_workflows(registry, manager)

        self.assertEqual([item["id"] for item in payload["workflows"]], ["hadith", "quran"])
        self.assertEqual(payload["statuses"]["hadith"]["state"], "idle")
```

- [ ] **Step 2: Run tests to verify failure**

Run: `python -m unittest tests.test_dashboard_server -v`

Expected: test fails because `api_workflows` is not defined.

- [ ] **Step 3: Implement API payload helper, request handler, and CLI**

Append to `dashboard_server.py`:

```python
def api_workflows(workflows: dict[str, Workflow], manager: JobManager) -> dict:
    return {
        "workflows": [workflows[key].to_dict() for key in ("hadith", "quran")],
        "statuses": {workflow_id: manager.status(workflow_id) for workflow_id in workflows},
    }


class DashboardRequestHandler(BaseHTTPRequestHandler):
    workflows: dict[str, Workflow]
    manager: JobManager

    def do_GET(self) -> None:
        parsed = urlparse(self.path)
        if parsed.path == "/api/workflows":
            self._send_json(api_workflows(self.workflows, self.manager))
            return
        if parsed.path.startswith("/api/status/"):
            workflow_id = parsed.path.rsplit("/", 1)[-1]
            self._send_json(self.manager.status(workflow_id))
            return
        self._serve_static(parsed.path)

    def do_POST(self) -> None:
        parsed = urlparse(self.path)
        if not parsed.path.startswith("/api/start/"):
            self._send_json({"error": "Not found"}, HTTPStatus.NOT_FOUND)
            return
        workflow_id = parsed.path.rsplit("/", 1)[-1]
        try:
            length = int(self.headers.get("Content-Length", "0"))
            body = self.rfile.read(length).decode("utf-8") if length else "{}"
            options = json.loads(body or "{}")
        except json.JSONDecodeError:
            self._send_json({"error": "Invalid JSON"}, HTTPStatus.BAD_REQUEST)
            return
        self._send_json(self.manager.start(workflow_id, options))

    def log_message(self, format: str, *args) -> None:
        return

    def _send_json(self, payload: dict, status: HTTPStatus = HTTPStatus.OK) -> None:
        data = json.dumps(payload, ensure_ascii=False).encode("utf-8")
        self.send_response(status)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(data)))
        self.end_headers()
        self.wfile.write(data)

    def _serve_static(self, request_path: str) -> None:
        relative = "index.html" if request_path in {"", "/"} else request_path.lstrip("/")
        file_path = (DASHBOARD_DIR / relative).resolve()
        try:
            file_path.relative_to(DASHBOARD_DIR.resolve())
        except ValueError:
            self.send_error(HTTPStatus.FORBIDDEN)
            return
        if not file_path.exists() or file_path.is_dir():
            self.send_error(HTTPStatus.NOT_FOUND)
            return
        content_type = mimetypes.guess_type(str(file_path))[0] or "application/octet-stream"
        data = file_path.read_bytes()
        self.send_response(HTTPStatus.OK)
        self.send_header("Content-Type", content_type)
        self.send_header("Content-Length", str(len(data)))
        self.end_headers()
        self.wfile.write(data)


def make_handler(workflows: dict[str, Workflow], manager: JobManager):
    class Handler(DashboardRequestHandler):
        pass

    Handler.workflows = workflows
    Handler.manager = manager
    return Handler


def main(argv: list[str] | None = None) -> int:
    import argparse
    import webbrowser

    parser = argparse.ArgumentParser(description="Run the local Hadith/Quran dashboard.")
    parser.add_argument("--host", default="127.0.0.1")
    parser.add_argument("--port", type=int, default=8765)
    parser.add_argument("--no-browser", action="store_true")
    args = parser.parse_args(argv)

    workflows = build_workflow_registry(ROOT_DIR)
    manager = JobManager(workflows)
    server = ThreadingHTTPServer((args.host, args.port), make_handler(workflows, manager))
    url = f"http://localhost:{args.port}" if args.host in {"127.0.0.1", "localhost"} else f"http://{args.host}:{args.port}"
    print(f"Dashboard running at {url}")
    if not args.no_browser:
        webbrowser.open(url)
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        print("\nDashboard stopped.")
    finally:
        server.server_close()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
```

- [ ] **Step 4: Run tests to verify pass**

Run: `python -m unittest tests.test_dashboard_server -v`

Expected: 7 tests pass.

## Task 4: Frontend Operator Console

**Files:**
- Create: `dashboard/index.html`
- Create: `dashboard/styles.css`
- Create: `dashboard/app.js`

- [ ] **Step 1: Create HTML shell**

Create `dashboard/index.html`:

```html
<!doctype html>
<html lang="en">
  <head>
    <meta charset="utf-8">
    <meta name="viewport" content="width=device-width, initial-scale=1">
    <title>Content Script Dashboard</title>
    <link rel="stylesheet" href="/styles.css">
  </head>
  <body>
    <main class="app-shell">
      <aside class="sidebar">
        <div class="brand">
          <span class="brand-mark">DB</span>
          <div>
            <h1>Content Dashboard</h1>
            <p>Local script console</p>
          </div>
        </div>
        <nav class="workflow-nav" id="workflowNav"></nav>
      </aside>

      <section class="workspace">
        <header class="workspace-header">
          <div>
            <p class="eyebrow">Selected workflow</p>
            <h2 id="workflowTitle">Loading...</h2>
            <p id="workflowDescription"></p>
          </div>
          <div class="status-pill" id="statusPill">Idle</div>
        </header>

        <section class="control-grid">
          <div class="panel controls-panel">
            <h3>Run Controls</h3>
            <label class="toggle-row" id="dryRunRow">
              <input type="checkbox" id="dryRunToggle" checked>
              <span>Dry run</span>
            </label>
            <button class="run-button" id="runButton" type="button">Run workflow</button>
          </div>

          <div class="panel paths-panel">
            <h3>Folders</h3>
            <div class="path-groups">
              <div>
                <h4>Inputs</h4>
                <ul id="inputPaths"></ul>
              </div>
              <div>
                <h4>Outputs</h4>
                <ul id="outputPaths"></ul>
              </div>
            </div>
          </div>
        </section>

        <section class="panel log-panel">
          <div class="log-header">
            <h3>Run Log</h3>
            <span id="exitCode"></span>
          </div>
          <pre id="logOutput">No run started yet.</pre>
        </section>
      </section>
    </main>
    <script src="/app.js"></script>
  </body>
</html>
```

- [ ] **Step 2: Create CSS**

Create `dashboard/styles.css` with a quiet operator-console layout:

```css
:root {
  color-scheme: light;
  --bg: #f4f6f8;
  --panel: #ffffff;
  --ink: #17202a;
  --muted: #627084;
  --line: #d8dee7;
  --accent: #0f766e;
  --accent-strong: #115e59;
  --danger: #b42318;
  --success: #067647;
  --warning: #b54708;
}

* {
  box-sizing: border-box;
}

body {
  margin: 0;
  font-family: Arial, Helvetica, sans-serif;
  background: var(--bg);
  color: var(--ink);
}

.app-shell {
  min-height: 100vh;
  display: grid;
  grid-template-columns: 280px 1fr;
}

.sidebar {
  background: #18232f;
  color: #f8fafc;
  padding: 22px 18px;
}

.brand {
  display: flex;
  gap: 12px;
  align-items: center;
  margin-bottom: 28px;
}

.brand-mark {
  width: 42px;
  height: 42px;
  display: inline-grid;
  place-items: center;
  border-radius: 8px;
  background: #0f766e;
  font-weight: 700;
}

.brand h1 {
  font-size: 18px;
  margin: 0;
}

.brand p,
#workflowDescription,
.eyebrow {
  margin: 4px 0 0;
  color: var(--muted);
}

.sidebar .brand p {
  color: #b7c3d0;
}

.workflow-nav {
  display: grid;
  gap: 8px;
}

.nav-button {
  width: 100%;
  min-height: 48px;
  border: 1px solid transparent;
  border-radius: 8px;
  background: transparent;
  color: #dce5ef;
  text-align: left;
  padding: 10px 12px;
  cursor: pointer;
}

.nav-button.active {
  background: #eef7f5;
  color: #12312f;
}

.workspace {
  padding: 28px;
  display: grid;
  gap: 18px;
  align-content: start;
}

.workspace-header,
.panel {
  background: var(--panel);
  border: 1px solid var(--line);
  border-radius: 8px;
}

.workspace-header {
  padding: 22px;
  display: flex;
  justify-content: space-between;
  gap: 16px;
  align-items: flex-start;
}

.eyebrow {
  text-transform: uppercase;
  font-size: 12px;
  letter-spacing: 0;
  font-weight: 700;
}

h2,
h3,
h4 {
  margin: 0;
}

h2 {
  font-size: 28px;
}

.status-pill {
  min-width: 108px;
  text-align: center;
  border-radius: 999px;
  border: 1px solid var(--line);
  padding: 8px 12px;
  font-weight: 700;
}

.status-pill.running {
  color: var(--warning);
  border-color: #fedf89;
  background: #fffaeb;
}

.status-pill.succeeded {
  color: var(--success);
  border-color: #abefc6;
  background: #ecfdf3;
}

.status-pill.failed {
  color: var(--danger);
  border-color: #fecdca;
  background: #fef3f2;
}

.control-grid {
  display: grid;
  grid-template-columns: 320px 1fr;
  gap: 18px;
}

.panel {
  padding: 18px;
}

.controls-panel {
  display: grid;
  gap: 16px;
  align-content: start;
}

.toggle-row {
  display: flex;
  align-items: center;
  gap: 10px;
  color: var(--ink);
}

.toggle-row input {
  width: 18px;
  height: 18px;
}

.run-button {
  min-height: 46px;
  border: 0;
  border-radius: 8px;
  background: var(--accent);
  color: white;
  font-weight: 700;
  cursor: pointer;
}

.run-button:disabled {
  opacity: 0.55;
  cursor: not-allowed;
}

.path-groups {
  display: grid;
  grid-template-columns: 1fr 1fr;
  gap: 18px;
}

ul {
  margin: 12px 0 0;
  padding: 0;
  list-style: none;
  display: grid;
  gap: 8px;
}

li {
  overflow-wrap: anywhere;
  border: 1px solid var(--line);
  border-radius: 6px;
  padding: 8px;
  background: #f8fafc;
  font-size: 13px;
}

.exists {
  border-left: 4px solid var(--success);
}

.missing {
  border-left: 4px solid var(--danger);
}

.log-panel {
  min-height: 340px;
}

.log-header {
  display: flex;
  justify-content: space-between;
  gap: 12px;
  margin-bottom: 12px;
}

pre {
  margin: 0;
  min-height: 260px;
  max-height: 460px;
  overflow: auto;
  white-space: pre-wrap;
  background: #111827;
  color: #e5e7eb;
  border-radius: 8px;
  padding: 14px;
  font-size: 13px;
  line-height: 1.5;
}

@media (max-width: 900px) {
  .app-shell,
  .control-grid,
  .path-groups {
    grid-template-columns: 1fr;
  }

  .workspace-header {
    flex-direction: column;
  }
}
```

- [ ] **Step 3: Create JavaScript app**

Create `dashboard/app.js`:

```javascript
const state = {
  workflows: [],
  statuses: {},
  selectedId: "hadith",
  pollTimer: null,
};

const workflowNav = document.getElementById("workflowNav");
const workflowTitle = document.getElementById("workflowTitle");
const workflowDescription = document.getElementById("workflowDescription");
const statusPill = document.getElementById("statusPill");
const dryRunRow = document.getElementById("dryRunRow");
const dryRunToggle = document.getElementById("dryRunToggle");
const runButton = document.getElementById("runButton");
const inputPaths = document.getElementById("inputPaths");
const outputPaths = document.getElementById("outputPaths");
const logOutput = document.getElementById("logOutput");
const exitCode = document.getElementById("exitCode");

async function fetchJson(url, options) {
  const response = await fetch(url, options);
  if (!response.ok) {
    throw new Error(`Request failed: ${response.status}`);
  }
  return response.json();
}

async function loadWorkflows() {
  const payload = await fetchJson("/api/workflows");
  state.workflows = payload.workflows;
  state.statuses = payload.statuses;
  render();
  startPolling();
}

function selectedWorkflow() {
  return state.workflows.find((workflow) => workflow.id === state.selectedId) || state.workflows[0];
}

function selectedStatus() {
  return state.statuses[state.selectedId] || { state: "idle", logs: [] };
}

function render() {
  const workflow = selectedWorkflow();
  if (!workflow) return;
  const status = selectedStatus();

  workflowNav.innerHTML = state.workflows.map((item) => `
    <button class="nav-button ${item.id === workflow.id ? "active" : ""}" type="button" data-id="${item.id}">
      ${item.name}
    </button>
  `).join("");

  workflowTitle.textContent = workflow.name;
  workflowDescription.textContent = workflow.description;
  statusPill.textContent = labelForState(status.state);
  statusPill.className = `status-pill ${status.state}`;
  dryRunRow.style.display = workflow.supports_dry_run ? "flex" : "none";
  runButton.disabled = status.state === "running";
  runButton.textContent = status.state === "running" ? "Running..." : `Run ${workflow.name}`;
  inputPaths.innerHTML = workflow.input_paths.map(renderPathItem).join("");
  outputPaths.innerHTML = (status.outputs || workflow.output_paths.map((path) => ({ path, exists: false }))).map(renderPathItem).join("");
  logOutput.textContent = status.logs && status.logs.length ? status.logs.join("\n") : "No run started yet.";
  exitCode.textContent = status.exit_code === null || status.exit_code === undefined ? "" : `Exit code: ${status.exit_code}`;

  document.querySelectorAll(".nav-button").forEach((button) => {
    button.addEventListener("click", () => {
      state.selectedId = button.dataset.id;
      render();
    });
  });
}

function renderPathItem(item) {
  const path = typeof item === "string" ? item : item.path;
  const exists = typeof item === "string" ? null : item.exists;
  const className = exists === null ? "" : exists ? "exists" : "missing";
  return `<li class="${className}">${path}</li>`;
}

function labelForState(value) {
  return {
    idle: "Idle",
    running: "Running",
    succeeded: "Succeeded",
    failed: "Failed",
  }[value] || value;
}

async function startWorkflow() {
  const workflow = selectedWorkflow();
  const options = { dry_run: Boolean(dryRunToggle.checked) };
  const status = await fetchJson(`/api/start/${workflow.id}`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(options),
  });
  state.statuses[workflow.id] = status;
  render();
}

function startPolling() {
  if (state.pollTimer) clearInterval(state.pollTimer);
  state.pollTimer = setInterval(refreshSelectedStatus, 1000);
}

async function refreshSelectedStatus() {
  if (!state.selectedId) return;
  const status = await fetchJson(`/api/status/${state.selectedId}`);
  state.statuses[state.selectedId] = status;
  render();
}

runButton.addEventListener("click", () => {
  startWorkflow().catch((error) => {
    logOutput.textContent = error.message;
  });
});

loadWorkflows().catch((error) => {
  workflowTitle.textContent = "Dashboard failed to load";
  workflowDescription.textContent = error.message;
});
```

- [ ] **Step 4: Smoke-test static files through server**

Run: `python -m unittest tests.test_dashboard_server -v`

Expected: existing server tests still pass.

## Task 5: Run and Verify Dashboard

**Files:**
- Modify only if verification exposes a defect.

- [ ] **Step 1: Run full unit tests**

Run: `python -m unittest discover -v`

Expected: dashboard tests pass. Existing script tests may require running from their script folders separately because they import local modules.

- [ ] **Step 2: Start dashboard server**

Run: `python dashboard_server.py --host 127.0.0.1 --port 8765 --no-browser`

Expected: console prints `Dashboard running at http://localhost:8765` and keeps running.

- [ ] **Step 3: Verify dashboard HTTP response**

In another shell, run: `Invoke-WebRequest -UseBasicParsing http://localhost:8765`

Expected: HTTP 200 and the response contains `Content Dashboard`.

- [ ] **Step 4: Give user dashboard URL**

Tell the user to open `http://localhost:8765`.
