from __future__ import annotations

import json
import os
import signal
import subprocess
import sys
import tempfile
import time
import unittest
from pathlib import Path

import openpyxl

try:
    import fcntl
except ImportError:  # pragma: no cover - non-POSIX test collection.
    fcntl = None

from formulabench.artifacts import (
    ArtifactError,
    ArtifactLayout,
    DatasetManifestError,
    ResumeStateError,
    append_jsonl,
    atomic_copy,
    atomic_write_json,
    confined_path,
    exclusive_output_lock,
    iter_jsonl,
    load_dataset_manifest,
    require_resume_directory,
    require_writable_directory,
    sha256_file,
)
from formulabench.capture import run_captured


def _write_workbook(path: Path, value: object = 1) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    workbook = openpyxl.Workbook()
    workbook.active["A1"] = value
    workbook.save(path)
    workbook.close()


class DatasetManifestTests(unittest.TestCase):
    def setUp(self) -> None:
        self.temporary = tempfile.TemporaryDirectory()
        self.root = Path(self.temporary.name)

    def tearDown(self) -> None:
        self.temporary.cleanup()

    def _manifest(self, tasks: list[dict]) -> None:
        (self.root / "dataset.json").write_text(json.dumps(tasks), encoding="utf-8")

    def _task(self, task_id: str | int = "51-12", folder: str = "spreadsheet/51-12") -> dict:
        return {
            "id": task_id,
            "instruction": "Fill the answer cell.",
            "spreadsheet_path": folder,
            "answer_position": "A1",
            "answer_sheet": "Sheet",
        }

    def test_loads_one_init_without_opening_golden_content(self) -> None:
        folder = self.root / "spreadsheet" / "51-12"
        _write_workbook(folder / "1_51-12_init.xlsx")
        # Deliberately not a workbook.  Loading succeeds because inference
        # discovery does not read or glob golden content.
        (folder / "1_51-12_golden.xlsx").write_bytes(b"not an OOXML workbook")
        self._manifest([self._task()])

        tasks = load_dataset_manifest(self.root)

        self.assertEqual([task.id for task in tasks], ["51-12"])
        self.assertEqual(tasks[0].init_xlsx.name, "1_51-12_init.xlsx")
        self.assertNotIn("golden_xlsx", tasks[0].as_dict())

    def test_loads_exact_initial_xlsx_without_opening_golden(self) -> None:
        folder = self.root / "spreadsheet" / "13284"
        _write_workbook(folder / "initial.xlsx")
        (folder / "golden.xlsx").write_bytes(b"not an OOXML workbook")
        self._manifest([self._task("13284", "spreadsheet/13284")])

        tasks = load_dataset_manifest(self.root)

        self.assertEqual(tasks[0].init_xlsx.name, "initial.xlsx")

    def test_rejects_mixed_initial_naming_conventions(self) -> None:
        folder = self.root / "spreadsheet" / "task"
        _write_workbook(folder / "1_task_init.xlsx")
        _write_workbook(folder / "initial.xlsx")
        self._manifest([self._task("task", "spreadsheet/task")])

        with self.assertRaises(DatasetManifestError):
            load_dataset_manifest(self.root)

    def test_rejects_duplicate_ids_after_normalisation(self) -> None:
        first = self.root / "spreadsheet" / "first"
        second = self.root / "spreadsheet" / "second"
        _write_workbook(first / "1_first_init.xlsx")
        _write_workbook(second / "1_second_init.xlsx")
        self._manifest(
            [
                self._task(7, "spreadsheet/first"),
                self._task("7", "spreadsheet/second"),
            ]
        )

        with self.assertRaises(DatasetManifestError):
            load_dataset_manifest(self.root)

    def test_rejects_case_colliding_ids_for_portable_output_paths(self) -> None:
        first = self.root / "spreadsheet" / "first"
        second = self.root / "spreadsheet" / "second"
        _write_workbook(first / "1_first_init.xlsx")
        _write_workbook(second / "1_second_init.xlsx")
        self._manifest(
            [
                self._task("Task-A", "spreadsheet/first"),
                self._task("task-a", "spreadsheet/second"),
            ]
        )

        with self.assertRaises(DatasetManifestError):
            load_dataset_manifest(self.root)

    def test_rejects_unsafe_ids(self) -> None:
        folder = self.root / "spreadsheet" / "task"
        _write_workbook(folder / "1_task_init.xlsx")
        for unsafe in ("../task", "task/name", "task\\name", ".", "", True, 1.5):
            with self.subTest(unsafe=unsafe):
                self._manifest([self._task(unsafe, "spreadsheet/task")])
                with self.assertRaises(DatasetManifestError):
                    load_dataset_manifest(self.root)

    def test_rejects_traversal_and_multiple_initial_workbooks(self) -> None:
        outside_temporary = tempfile.TemporaryDirectory(dir=self.root.parent)
        self.addCleanup(outside_temporary.cleanup)
        outside = Path(outside_temporary.name)
        _write_workbook(outside / "1_outside_init.xlsx")
        self._manifest([self._task("task", f"../{outside.name}")])
        with self.assertRaises(DatasetManifestError):
            load_dataset_manifest(self.root)

        folder = self.root / "spreadsheet" / "task"
        _write_workbook(folder / "1_task_init.xlsx")
        _write_workbook(folder / "2_task_init.xlsx")
        self._manifest([self._task("task", "spreadsheet/task")])
        with self.assertRaises(DatasetManifestError):
            load_dataset_manifest(self.root)

    @unittest.skipUnless(hasattr(os, "symlink"), "symbolic links are unavailable")
    def test_rejects_symlink_escape_for_task_folder_and_init(self) -> None:
        outside_temporary = tempfile.TemporaryDirectory(dir=self.root.parent)
        self.addCleanup(outside_temporary.cleanup)
        outside = Path(outside_temporary.name)
        _write_workbook(outside / "1_external_init.xlsx")

        spreadsheet = self.root / "spreadsheet"
        spreadsheet.mkdir()
        try:
            (spreadsheet / "linked").symlink_to(outside, target_is_directory=True)
        except OSError as exc:  # pragma: no cover - restricted Windows accounts.
            self.skipTest(str(exc))
        self._manifest([self._task("linked", "spreadsheet/linked")])
        with self.assertRaises(DatasetManifestError):
            load_dataset_manifest(self.root)

        (spreadsheet / "linked").unlink()
        folder = spreadsheet / "linked"
        folder.mkdir()
        (folder / "1_linked_init.xlsx").symlink_to(outside / "1_external_init.xlsx")
        with self.assertRaises(DatasetManifestError):
            load_dataset_manifest(self.root)


class ArtifactHelperTests(unittest.TestCase):
    def setUp(self) -> None:
        self.temporary = tempfile.TemporaryDirectory()
        self.root = Path(self.temporary.name)

    def tearDown(self) -> None:
        self.temporary.cleanup()

    def test_confined_path_rejects_traversal(self) -> None:
        (self.root / "inside.txt").write_text("inside", encoding="utf-8")
        with self.assertRaises(ArtifactError):
            confined_path(self.root, "../outside.txt")

    def test_non_empty_output_is_preserved(self) -> None:
        out = self.root / "out"
        out.mkdir()
        sentinel = out / "keep.txt"
        sentinel.write_text("keep", encoding="utf-8")

        with self.assertRaises(ArtifactError):
            require_writable_directory(out)

        self.assertEqual(sentinel.read_text(encoding="utf-8"), "keep")
        self.assertEqual([item.name for item in out.iterdir()], ["keep.txt"])

    def test_atomic_jsonl_copy_and_hash(self) -> None:
        source = self.root / "source.bin"
        source.write_bytes(os.urandom(4096))
        destination = self.root / "nested" / "copy.bin"
        atomic_copy(source, destination)
        self.assertEqual(sha256_file(source), sha256_file(destination))

        records = self.root / "records.jsonl"
        append_jsonl(records, {"id": "a", "value": "line\nbreak"})
        append_jsonl(records, {"id": "b", "value": 2})
        self.assertEqual(
            [record for _, record in iter_jsonl(records)],
            [
                {"id": "a", "value": "line\nbreak"},
                {"id": "b", "value": 2},
            ],
        )

        document = self.root / "document.json"
        atomic_write_json(document, {"b": 2, "a": 1})
        self.assertEqual(json.loads(document.read_text(encoding="utf-8")), {"a": 1, "b": 2})

    def test_layout_allows_only_the_supervisor_log(self) -> None:
        out = self.root / "out"
        out.mkdir()
        (out / "run.log").write_bytes(b"")

        layout = ArtifactLayout.initialise(out, allow_supervisor_log=True)

        self.assertTrue(layout.outputs.is_dir())
        self.assertTrue(layout.traces.is_dir())
        self.assertEqual(layout.predictions.read_bytes(), b"")

    def test_resume_layout_completes_only_missing_empty_structure(self) -> None:
        out = self.root / "out"
        out.mkdir()
        original_log = b"first attempt\x00\n"
        (out / "run.log").write_bytes(original_log)

        layout = ArtifactLayout.resume(out)

        self.assertEqual(layout.run_log.read_bytes(), original_log)
        self.assertEqual(layout.predictions.read_bytes(), b"")
        self.assertTrue(layout.outputs.is_dir())
        self.assertTrue(layout.traces.is_dir())
        self.assertEqual(require_resume_directory(out), out.resolve())

    def test_resume_layout_rejects_and_preserves_unexpected_user_data(self) -> None:
        out = self.root / "out"
        out.mkdir()
        (out / "run.log").write_bytes(b"old\n")
        sentinel = out / "notes.txt"
        sentinel.write_text("preserve", encoding="utf-8")

        with self.assertRaisesRegex(ResumeStateError, "unexpected_root_entry"):
            ArtifactLayout.resume(out)

        self.assertEqual(sentinel.read_text(encoding="utf-8"), "preserve")

    def test_capture_preserves_combined_raw_output(self) -> None:
        out = self.root / "captured"
        code = "import os; os.write(1,b'out\\x00\\n'); os.write(2,b'err\\xff\\n')"

        result = run_captured([sys.executable, "-c", code], out, echo=False)

        self.assertEqual(result.returncode, 0)
        self.assertEqual(result.run_log.read_bytes(), b"out\x00\nerr\xff\n")
        self.assertEqual([path.name for path in out.iterdir()], ["run.log"])

    @unittest.skipIf(fcntl is None, "advisory file locks require POSIX")
    def test_capture_passes_its_output_lock_to_the_child(self) -> None:
        out = self.root / "captured"
        code = (
            "from pathlib import Path\n"
            "from formulabench.artifacts import exclusive_output_lock\n"
            f"with exclusive_output_lock(Path({str(out / 'run.log')!r})):\n"
            "    print('locked')\n"
        )

        result = run_captured([sys.executable, "-c", code], out, echo=False)

        self.assertEqual(result.returncode, 0)
        self.assertEqual(result.run_log.read_text(encoding="utf-8"), "locked\n")

    @unittest.skipIf(fcntl is None, "advisory file locks require POSIX")
    def test_direct_output_lock_rejects_a_concurrent_owner(self) -> None:
        out = self.root / "out"
        out.mkdir()
        run_log = out / "run.log"
        run_log.write_bytes(b"")
        lock_fd = os.open(run_log, os.O_RDWR)
        self.addCleanup(os.close, lock_fd)
        assert fcntl is not None
        fcntl.flock(lock_fd, fcntl.LOCK_EX | fcntl.LOCK_NB)

        with (
            self.assertRaisesRegex(ResumeStateError, "run_already_active"),
            exclusive_output_lock(run_log),
        ):
            self.fail("concurrent owner acquired the output lock")

    def test_capture_resume_appends_without_changing_existing_log_bytes(self) -> None:
        out = self.root / "captured"
        first = b"first\x00\xff\n"
        second = b"second\xfe\n"
        run_captured(
            [sys.executable, "-c", f"import os; os.write(1,{first!r})"],
            out,
            echo=False,
        )
        before = (out / "run.log").read_bytes()

        result = run_captured(
            [sys.executable, "-c", f"import os; os.write(2,{second!r})"],
            out,
            echo=False,
            resume=True,
        )

        self.assertEqual(result.returncode, 0)
        self.assertEqual(before, first)
        self.assertEqual(result.run_log.read_bytes(), first + second)

    @unittest.skipIf(fcntl is None, "advisory file locks require POSIX")
    def test_capture_resume_rejects_a_concurrent_run_without_touching_log(self) -> None:
        out = self.root / "captured"
        out.mkdir()
        run_log = out / "run.log"
        original = b"active run\n"
        run_log.write_bytes(original)
        lock_fd = os.open(run_log, os.O_WRONLY | os.O_APPEND)
        self.addCleanup(os.close, lock_fd)
        assert fcntl is not None
        fcntl.flock(lock_fd, fcntl.LOCK_EX | fcntl.LOCK_NB)

        with self.assertRaisesRegex(ResumeStateError, "run_already_active"):
            run_captured(
                [sys.executable, "-c", "print('must not run')"],
                out,
                echo=False,
                resume=True,
            )

        self.assertEqual(run_log.read_bytes(), original)

    def test_capture_never_cleans_existing_output(self) -> None:
        out = self.root / "captured"
        out.mkdir()
        sentinel = out / "existing.xlsx"
        sentinel.write_bytes(b"preserve")

        with self.assertRaises(ArtifactError):
            run_captured([sys.executable, "-c", "print('should not run')"], out, echo=False)

        self.assertEqual(sentinel.read_bytes(), b"preserve")
        self.assertFalse((out / "run.log").exists())

    @unittest.skipUnless(os.name == "posix", "process-group signalling requires POSIX")
    def test_capture_forwards_termination_signal_to_child_group(self) -> None:
        out = self.root / "signalled"
        child_code = (
            "import os,signal,time; "
            "signal.signal(signal.SIGTERM, "
            "lambda *_: (os.write(1,b'forwarded\\n'), exit(0))); "
            "os.write(1,b'ready\\n'); time.sleep(30)"
        )
        supervisor_code = (
            "import sys; from formulabench.capture import run_captured; "
            f"result=run_captured([sys.executable,'-c',{child_code!r}],{str(out)!r},echo=False); "
            "raise SystemExit(result.returncode)"
        )
        supervisor = subprocess.Popen([sys.executable, "-c", supervisor_code])
        self.addCleanup(lambda: supervisor.kill() if supervisor.poll() is None else None)

        deadline = time.monotonic() + 5
        while time.monotonic() < deadline:
            log = out / "run.log"
            if log.exists() and b"ready\n" in log.read_bytes():
                break
            time.sleep(0.02)
        else:
            self.fail("child did not become ready")

        supervisor.send_signal(signal.SIGTERM)
        self.assertEqual(supervisor.wait(timeout=5), 0)
        self.assertEqual((out / "run.log").read_bytes(), b"ready\nforwarded\n")


if __name__ == "__main__":
    unittest.main()
