import re
import sys
import tempfile
import threading
import unittest
from pathlib import Path

import dashboard_server as dashboard


class WorkflowRegistryTest(unittest.TestCase):
    def test_registry_contains_hadith_and_quran_workflows(self):
        registry = dashboard.build_workflow_registry(Path.cwd())

        self.assertEqual(set(registry), {"hadith", "quran", "quran-update-json", "dua", "dua-update-json"})
        self.assertEqual(registry["hadith"].name, "Single Hadith Content")
        self.assertEqual(registry["quran"].name, "Quran Content")
        self.assertEqual(registry["quran-update-json"].name, "Update Quran Output JSON")
        self.assertEqual(registry["dua"].name, "Dua Content")
        self.assertEqual(registry["dua-update-json"].name, "Update Dua Output JSON")
        self.assertEqual(registry["hadith"].working_dir.name, "SINGLE HADITH CONTENT")
        self.assertEqual(registry["quran"].working_dir.name, "QURAN CONTENT")
        self.assertEqual(registry["dua"].working_dir.name, "DUA CONTENT")

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

    def test_quran_update_json_command_includes_structure_and_files(self):
        registry = dashboard.build_workflow_registry(Path.cwd())

        command = registry["quran-update-json"].command(
            {"structure": "surah.json", "files": ["EN_Quran Surah Content.xlsx", "EN_Quran Page Content.xlsx"]}
        )

        self.assertEqual(
            command,
            [
                sys.executable,
                "convert_quran_content.py",
                "update-output-json",
                "--structure",
                "surah.json",
                "--files",
                "EN_Quran Surah Content.xlsx",
                "EN_Quran Page Content.xlsx",
            ],
        )

    def test_dua_commands_include_convert_and_selected_update_files(self):
        registry = dashboard.build_workflow_registry(Path.cwd())

        convert_command = registry["dua"].command({})
        update_command = registry["dua-update-json"].command({"files": ["EN _SINGLE DUA.xlsx", "BN_SINGLE DUA.xlsx"]})

        self.assertEqual(convert_command, [sys.executable, "convert_dua_content.py", "convert"])
        self.assertEqual(
            update_command,
            [
                sys.executable,
                "convert_dua_content.py",
                "update-output-json",
                "--files",
                "EN _SINGLE DUA.xlsx",
                "BN_SINGLE DUA.xlsx",
            ],
        )


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


class ApiResponseTest(unittest.TestCase):
    def test_api_workflows_returns_serializable_workflow_list(self):
        registry = dashboard.build_workflow_registry(Path.cwd())
        manager = dashboard.JobManager(registry)

        payload = dashboard.api_workflows(registry, manager)

        self.assertEqual([item["id"] for item in payload["workflows"]], ["hadith", "quran", "dua"])
        self.assertEqual(payload["statuses"]["hadith"]["state"], "idle")

    def test_api_quran_output_files_lists_xlsx_files(self):
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            output_dir = root / "QURAN CONTENT" / "OUTPUT"
            output_dir.mkdir(parents=True)
            (output_dir / "EN_Quran Juz Content.xlsx").touch()
            (output_dir / "~$temporary.xlsx").touch()
            (output_dir / "conversion_summary.csv").touch()

            payload = dashboard.api_quran_output_files(root)

        self.assertEqual(payload, {"files": ["EN_Quran Juz Content.xlsx"]})

    def test_api_dua_output_files_lists_xlsx_files(self):
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            output_dir = root / "DUA CONTENT" / "OUTPUT"
            output_dir.mkdir(parents=True)
            (output_dir / "EN _SINGLE DUA.xlsx").touch()
            (output_dir / "~$temporary.xlsx").touch()
            (output_dir / "conversion_summary.csv").touch()

            payload = dashboard.api_dua_output_files(root)

        self.assertEqual(payload, {"files": ["EN _SINGLE DUA.xlsx"]})


class DashboardUiTest(unittest.TestCase):
    def test_dry_run_toggle_is_not_enabled_by_default(self):
        html = (Path(__file__).resolve().parents[1] / "dashboard" / "index.html").read_text(encoding="utf-8")

        match = re.search(r"<input[^>]+id=\"dryRunToggle\"[^>]*>", html)

        self.assertIsNotNone(match)
        self.assertNotRegex(match.group(0), r"\bchecked\b")
