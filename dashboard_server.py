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
    command_builder: Callable[[dict], list[str]] | None = None

    def command(self, options: dict | None = None) -> list[str]:
        options = options or {}
        if self.command_builder is not None:
            return self.command_builder(options)
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
    dua_dir = root_dir / "DUA CONTENT"
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
        "hadith-book-wise-final": Workflow(
            id="hadith-book-wise-final",
            name="Build Book Wise Final",
            description="Parse selected reconciled_output workbooks into book-wise final workbooks.",
            script="reconcile_hadith_outputs.py",
            working_dir=hadith_dir,
            input_paths=(hadith_dir / "reconciled_output", hadith_dir / "BOOKS.xlsx"),
            output_paths=(hadith_dir / "BOOK WISE FINAL",),
            command_builder=_hadith_book_wise_final_command,
        ),
        "hadith-updated-content": Workflow(
            id="hadith-updated-content",
            name="Update Book Wise Final Content",
            description="Parse selected book-wise final workbooks into updated content workbooks.",
            script="reconcile_hadith_outputs.py",
            working_dir=hadith_dir,
            input_paths=(hadith_dir / "BOOK WISE FINAL", hadith_dir / "BOOKS.xlsx"),
            output_paths=(hadith_dir / "BOOK WISE FINAL UPDATED CONTENT",),
            command_builder=_hadith_updated_content_command,
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
        "quran-update-json": Workflow(
            id="quran-update-json",
            name="Update Quran Output JSON",
            description="Rewrite selected Quran OUTPUT workbook contents into the selected JSON structure.",
            script="convert_quran_content.py",
            working_dir=quran_dir,
            input_paths=(quran_dir / "OUTPUT",),
            output_paths=(quran_dir / "UPDATED OUTPUT",),
            command_builder=_quran_update_json_command,
        ),
        "dua": Workflow(
            id="dua",
            name="Dua Content",
            description="Convert Dua input workbooks to id, language_id, and contents columns.",
            script="convert_dua_content.py",
            working_dir=dua_dir,
            input_paths=(dua_dir / "INPUT",),
            output_paths=(dua_dir / "OUTPUT", dua_dir / "OUTPUT" / "conversion_summary.csv"),
            command_builder=lambda _options: [sys.executable, "convert_dua_content.py", "convert"],
        ),
        "dua-update-json": Workflow(
            id="dua-update-json",
            name="Update Dua Output JSON",
            description="Rewrite selected Dua OUTPUT workbook contents into the approved JSON structure.",
            script="convert_dua_content.py",
            working_dir=dua_dir,
            input_paths=(dua_dir / "OUTPUT",),
            output_paths=(dua_dir / "UPDATED CONTENT",),
            command_builder=_dua_update_json_command,
        ),
    }


def _hadith_book_wise_final_command(options: dict) -> list[str]:
    files = options.get("files") or []
    command = [
        sys.executable,
        "reconcile_hadith_outputs.py",
        "write-book-wise-final",
    ]
    if files:
        command.append("--files")
        command.extend(str(file_name) for file_name in files)
    return command


def _hadith_updated_content_command(options: dict) -> list[str]:
    files = options.get("files") or []
    command = [
        sys.executable,
        "reconcile_hadith_outputs.py",
        "update-final-content",
    ]
    if files:
        command.append("--files")
        command.extend(str(file_name) for file_name in files)
    return command


def _quran_update_json_command(options: dict) -> list[str]:
    structure = options.get("structure") or "juz.json"
    files = options.get("files") or []
    command = [
        sys.executable,
        "convert_quran_content.py",
        "update-output-json",
        "--structure",
        structure,
    ]
    if files:
        command.append("--files")
        command.extend(str(file_name) for file_name in files)
    return command


def _dua_update_json_command(options: dict) -> list[str]:
    files = options.get("files") or []
    command = [
        sys.executable,
        "convert_dua_content.py",
        "update-output-json",
    ]
    if files:
        command.append("--files")
        command.extend(str(file_name) for file_name in files)
    return command


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


def api_workflows(workflows: dict[str, Workflow], manager: JobManager) -> dict:
    return {
        "workflows": [workflows[key].to_dict() for key in ("hadith", "quran", "dua")],
        "statuses": {workflow_id: manager.status(workflow_id) for workflow_id in workflows},
    }


def api_hadith_reconciled_files(root_dir: Path = ROOT_DIR) -> dict:
    output_dir = root_dir / "SINGLE HADITH CONTENT" / "reconciled_output"
    return {"files": _relative_xlsx_files(output_dir)}


def api_hadith_book_wise_final_files(root_dir: Path = ROOT_DIR) -> dict:
    final_dir = root_dir / "SINGLE HADITH CONTENT" / "BOOK WISE FINAL"
    return {"files": _relative_xlsx_files(final_dir)}


def api_quran_output_files(root_dir: Path = ROOT_DIR) -> dict:
    output_dir = root_dir / "QURAN CONTENT" / "OUTPUT"
    return {"files": _relative_xlsx_files(output_dir)}


def api_dua_output_files(root_dir: Path = ROOT_DIR) -> dict:
    output_dir = root_dir / "DUA CONTENT" / "OUTPUT"
    return {"files": _relative_xlsx_files(output_dir)}


def _relative_xlsx_files(directory: Path) -> list[str]:
    if not directory.exists():
        return []
    return sorted(
        path.relative_to(directory).as_posix()
        for path in directory.rglob("*.xlsx")
        if not path.name.startswith("~$")
    )


class DashboardRequestHandler(BaseHTTPRequestHandler):
    workflows: dict[str, Workflow]
    manager: JobManager

    def do_GET(self) -> None:
        parsed = urlparse(self.path)
        if parsed.path == "/api/workflows":
            self._send_json(api_workflows(self.workflows, self.manager))
            return
        if parsed.path == "/api/hadith-reconciled-files":
            self._send_json(api_hadith_reconciled_files(ROOT_DIR))
            return
        if parsed.path == "/api/hadith-book-wise-final-files":
            self._send_json(api_hadith_book_wise_final_files(ROOT_DIR))
            return
        if parsed.path == "/api/quran-output-files":
            self._send_json(api_quran_output_files(ROOT_DIR))
            return
        if parsed.path == "/api/dua-output-files":
            self._send_json(api_dua_output_files(ROOT_DIR))
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
