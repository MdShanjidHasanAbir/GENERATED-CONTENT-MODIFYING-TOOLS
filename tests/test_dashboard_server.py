import re
import sys
import tempfile
import threading
import unittest
import base64
from pathlib import Path

import dashboard_server as dashboard


class WorkflowRegistryTest(unittest.TestCase):
    def test_registry_contains_hadith_and_quran_workflows(self):
        registry = dashboard.build_workflow_registry(Path.cwd())

        self.assertEqual(
            set(registry),
            {
                "hadith",
                "hadith-book-wise-final",
                "hadith-updated-content",
                "quran",
                "quran-update-json",
                "dua",
                "dua-update-json",
            },
        )
        self.assertEqual(registry["hadith"].name, "Single Hadith Content")
        self.assertEqual(registry["hadith-book-wise-final"].name, "Build Book Wise Final")
        self.assertEqual(registry["hadith-updated-content"].name, "Update Book Wise Final Content")
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

    def test_hadith_stage_commands_include_selected_files(self):
        registry = dashboard.build_workflow_registry(Path.cwd())

        final_command = registry["hadith-book-wise-final"].command(
            {"files": ["EN_BUKHARI_MARGE.xlsx", "BN_MUSLIM_MARGE.xlsx"]}
        )
        updated_command = registry["hadith-updated-content"].command(
            {"files": ["EN/EN_BUKHARI.xlsx", "BN/BN_MUSLIM.xlsx"]}
        )

        self.assertEqual(
            final_command,
            [
                sys.executable,
                "reconcile_hadith_outputs.py",
                "write-book-wise-final",
                "--files",
                "EN_BUKHARI_MARGE.xlsx",
                "BN_MUSLIM_MARGE.xlsx",
            ],
        )
        self.assertEqual(
            updated_command,
            [
                sys.executable,
                "reconcile_hadith_outputs.py",
                "update-final-content",
                "--files",
                "EN/EN_BUKHARI.xlsx",
                "BN/BN_MUSLIM.xlsx",
            ],
        )

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
        self.assertEqual(payload["statuses"]["hadith-book-wise-final"]["state"], "idle")

    def test_api_hadith_reconciled_files_lists_xlsx_files(self):
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            output_dir = root / "SINGLE HADITH CONTENT" / "reconciled_output"
            output_dir.mkdir(parents=True)
            (output_dir / "EN_BUKHARI_MARGE.xlsx").touch()
            (output_dir / "~$temporary.xlsx").touch()
            (output_dir / "reconciliation_summary.csv").touch()

            payload = dashboard.api_hadith_reconciled_files(root)

        self.assertEqual(payload, {"files": ["EN_BUKHARI_MARGE.xlsx"]})

    def test_api_hadith_book_wise_final_files_lists_relative_xlsx_files(self):
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            final_dir = root / "SINGLE HADITH CONTENT" / "BOOK WISE FINAL"
            (final_dir / "EN").mkdir(parents=True)
            (final_dir / "BN").mkdir(parents=True)
            (final_dir / "EN" / "EN_BUKHARI.xlsx").touch()
            (final_dir / "BN" / "BN_MUSLIM.xlsx").touch()
            (final_dir / "EN" / "~$temporary.xlsx").touch()

            payload = dashboard.api_hadith_book_wise_final_files(root)

        self.assertEqual(payload, {"files": ["BN/BN_MUSLIM.xlsx", "EN/EN_BUKHARI.xlsx"]})

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

    def test_upload_target_files_use_approved_folders(self):
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            (root / "QURAN CONTENT" / "INPUT").mkdir(parents=True)
            (root / "DUA CONTENT" / "INPUT").mkdir(parents=True)
            (root / "SINGLE HADITH CONTENT" / "OUTPUT").mkdir(parents=True)
            (root / "QURAN CONTENT" / "INPUT" / "Quran.xlsx").touch()
            (root / "DUA CONTENT" / "INPUT" / "Dua.xlsx").touch()
            (root / "SINGLE HADITH CONTENT" / "OUTPUT" / "Hadith.xlsx").touch()

            self.assertEqual(dashboard.api_upload_target_files("quran", root), {"files": ["Quran.xlsx"]})
            self.assertEqual(dashboard.api_upload_target_files("dua", root), {"files": ["Dua.xlsx"]})
            self.assertEqual(dashboard.api_upload_target_files("hadith", root), {"files": ["Hadith.xlsx"]})

    def test_save_uploaded_files_writes_xlsx_to_target_and_rejects_unsafe_names(self):
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            content = base64.b64encode(b"excel bytes").decode("ascii")

            payload = dashboard.save_uploaded_files(
                "quran",
                {
                    "overwrite": False,
                    "files": [
                        {"name": "Quran.xlsx", "content": content},
                        {"name": "~$Quran.xlsx", "content": content},
                        {"name": "../outside.xlsx", "content": content},
                        {"name": "notes.txt", "content": content},
                    ],
                },
                root,
            )

            target = root / "QURAN CONTENT" / "INPUT" / "Quran.xlsx"
            self.assertEqual(target.read_bytes(), b"excel bytes")
            self.assertEqual(payload["saved"], ["Quran.xlsx"])
            self.assertEqual(
                payload["skipped"],
                [
                    {"name": "~$Quran.xlsx", "reason": "Temporary Excel files are not allowed."},
                    {"name": "../outside.xlsx", "reason": "Folder paths are not allowed."},
                    {"name": "notes.txt", "reason": "Only .xlsx files are allowed."},
                ],
            )
            self.assertFalse((root / "outside.xlsx").exists())

    def test_save_uploaded_files_skips_existing_without_overwrite(self):
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            target_dir = root / "DUA CONTENT" / "INPUT"
            target_dir.mkdir(parents=True)
            (target_dir / "Dua.xlsx").write_bytes(b"old")
            content = base64.b64encode(b"new").decode("ascii")

            payload = dashboard.save_uploaded_files(
                "dua",
                {"overwrite": False, "files": [{"name": "Dua.xlsx", "content": content}]},
                root,
            )

            self.assertEqual((target_dir / "Dua.xlsx").read_bytes(), b"old")
            self.assertEqual(payload["saved"], [])
            self.assertEqual(payload["skipped"], [{"name": "Dua.xlsx", "reason": "File already exists."}])


class DashboardUiTest(unittest.TestCase):
    def test_dry_run_toggle_is_not_enabled_by_default(self):
        html = (Path(__file__).resolve().parents[1] / "dashboard" / "index.html").read_text(encoding="utf-8")

        match = re.search(r"<input[^>]+id=\"dryRunToggle\"[^>]*>", html)

        self.assertIsNotNone(match)
        self.assertNotRegex(match.group(0), r"\bchecked\b")

    def test_folder_path_panel_is_not_rendered(self):
        html = (Path(__file__).resolve().parents[1] / "dashboard" / "index.html").read_text(encoding="utf-8")

        self.assertNotIn('id="inputPaths"', html)
        self.assertNotIn('id="outputPaths"', html)

    def test_upload_controls_exist_for_main_workflows(self):
        html = (Path(__file__).resolve().parents[1] / "dashboard" / "index.html").read_text(encoding="utf-8")

        self.assertIn('id="hadithUploadInput"', html)
        self.assertIn('id="quranUploadInput"', html)
        self.assertIn('id="duaUploadInput"', html)
