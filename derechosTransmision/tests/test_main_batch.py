import sys
import types
import unittest
from pathlib import Path
from unittest.mock import patch


SRC_DIR = Path(__file__).resolve().parents[1] / "src"
if str(SRC_DIR) not in sys.path:
    sys.path.insert(0, str(SRC_DIR))

googleapiclient_module = sys.modules.setdefault("googleapiclient", types.ModuleType("googleapiclient"))
googleapiclient_module.__path__ = []
googleapiclient_errors = sys.modules.setdefault(
    "googleapiclient.errors", types.ModuleType("googleapiclient.errors")
)


class HttpError(Exception):
    pass


googleapiclient_errors.HttpError = HttpError
googleapiclient_module.errors = googleapiclient_errors

if "drive_service" not in sys.modules:
    drive_service_stub = types.ModuleType("drive_service")
    drive_service_stub.get_drive_service = lambda *args, **kwargs: object()
    drive_service_stub.list_child_folders = lambda *args, **kwargs: []
    drive_service_stub.download_file_from_drive = lambda *args, **kwargs: None
    drive_service_stub.upload_file_to_folder = lambda *args, **kwargs: None
    drive_service_stub.validate_required_folder_files = lambda *args, **kwargs: None
    sys.modules["drive_service"] = drive_service_stub

if "pdf_service" not in sys.modules:
    pdf_service_stub = types.ModuleType("pdf_service")
    pdf_service_stub.convert_docx_to_pdf = lambda *args, **kwargs: None
    pdf_service_stub.generate_acta_docx = lambda *args, **kwargs: None
    pdf_service_stub.generate_anexo_qrs = lambda *args, **kwargs: None
    pdf_service_stub.generate_relacion_anexos_docx = lambda *args, **kwargs: None
    pdf_service_stub.generate_relacion_anexos_qr = lambda *args, **kwargs: None
    pdf_service_stub.read_acta_data_xlsx = lambda *args, **kwargs: None
    sys.modules["pdf_service"] = pdf_service_stub

import main  # noqa: E402


class MainBatchTests(unittest.TestCase):
    def test_normalize_mode_supports_single_and_batch_aliases(self):
        self.assertEqual(main._normalize_mode("1"), "single")
        self.assertEqual(main._normalize_mode("single"), "single")
        self.assertEqual(main._normalize_mode("2"), "batch")
        self.assertEqual(main._normalize_mode("batch"), "batch")
        self.assertIsNone(main._normalize_mode("unknown"))

    def test_run_multiple_folders_continues_after_validation_failure_and_exception(self):
        child_folders = [
            {"id": "folder-1", "name": "Folder Uno", "link": "https://drive.google.com/drive/folders/folder-1"},
            {"id": "folder-2", "name": "Folder Dos", "link": "https://drive.google.com/drive/folders/folder-2"},
            {"id": "folder-3", "name": "Folder Tres", "link": "https://drive.google.com/drive/folders/folder-3"},
        ]
        service = object()
        run_results = [True, False, RuntimeError("boom")]

        def fake_run_single_folder(folder_link, service=None):
            outcome = run_results.pop(0)
            if isinstance(outcome, Exception):
                raise outcome
            return outcome

        with patch.object(main, "get_drive_service", return_value=service), \
             patch.object(main, "list_child_folders", return_value=child_folders), \
             patch.object(main, "_run_single_folder", side_effect=fake_run_single_folder), \
             patch.object(main.traceback, "print_exc"), \
             patch("builtins.print") as print_mock:
            main._run_multiple_folders("https://drive.google.com/drive/folders/parent")

        printed_lines = [" ".join(str(arg) for arg in call.args) for call in print_mock.call_args_list]
        self.assertTrue(any("Found 3 child folder(s) to process." in line for line in printed_lines))
        self.assertTrue(any("Successful folders: 1" in line for line in printed_lines))
        self.assertTrue(any("Failed folders: 2" in line for line in printed_lines))
        self.assertTrue(any("Folder Dos: Validation failed" in line for line in printed_lines))
        self.assertTrue(any("Folder Tres: RuntimeError: boom" in line for line in printed_lines))

    def test_main_single_mode_uses_single_folder_prompt(self):
        with patch("builtins.input", side_effect=["1", "https://drive.google.com/drive/folders/abc"]), \
             patch.object(main, "_run_single_folder") as run_single_mock:
            main.main()

        run_single_mock.assert_called_once_with("https://drive.google.com/drive/folders/abc")


if __name__ == "__main__":
    unittest.main()
