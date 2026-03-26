import sys
import types
import unittest
from pathlib import Path
from unittest.mock import patch
import shutil
import tempfile


SRC_DIR = Path(__file__).resolve().parents[1] / "src"
if str(SRC_DIR) not in sys.path:
    sys.path.insert(0, str(SRC_DIR))

if "drive_service" not in sys.modules:
    drive_service_stub = types.ModuleType("drive_service")
    drive_service_stub.download_file_from_drive = lambda *args, **kwargs: None
    drive_service_stub.get_drive_service = lambda *args, **kwargs: None
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

import pipeline  # noqa: E402


class PipelineCleanupTests(unittest.TestCase):
    def setUp(self):
        self.runs_path = Path(__file__).resolve().parents[1] / "runs"
        self.validation_result = {
            "folder_id": "folder-123",
            "folder_name": "Restaurant Norte",
            "links": {
                "anexoA": "https://example.com/anexoA",
                "anexoB": "https://example.com/anexoB",
                "anexoC": "https://example.com/anexoC",
            },
            "ids": {
                "acta_data": "file-acta",
            },
        }
        self.created_run_folders = []
        self.addCleanup(self._cleanup_created_run_folders)

    def _cleanup_created_run_folders(self):
        for folder in self.created_run_folders:
            if folder.exists():
                shutil.rmtree(folder, ignore_errors=True)

    def _fake_create_run_folder(self, folder_name: str) -> Path:
        folder = self.runs_path / f"{folder_name.replace(' ', '_')}_test_run_{len(self.created_run_folders)}"
        folder.mkdir(parents=True, exist_ok=True)
        self.created_run_folders.append(folder)
        return folder

    def _write_file(self, path: Path, contents: str = "data") -> str:
        path.parent.mkdir(parents=True, exist_ok=True)
        path.write_text(contents, encoding="utf-8")
        return str(path)

    def test_validation_failure_does_not_create_run_folder(self):
        with patch.object(pipeline, "get_drive_service", return_value=object()), \
             patch.object(pipeline, "validate_required_folder_files", return_value=None), \
             patch.object(pipeline, "_create_run_folder") as create_run_folder_mock, \
             patch.object(pipeline, "download_file_from_drive") as download_mock:
            result = pipeline.run_pipeline("https://drive.google.com/folders/test")

        self.assertIsNone(result)
        create_run_folder_mock.assert_not_called()
        download_mock.assert_not_called()

    def test_failed_run_deletes_temp_folder(self):
        def fake_download_file_from_drive(service, file_id, destination_path):
            return self._write_file(Path(destination_path), "xlsx")

        def fail_generating_qrs(validation_result, output_folder):
            raise RuntimeError("QR generation exploded")

        with patch.object(pipeline, "get_drive_service", return_value=object()), \
             patch.object(
                 pipeline,
                 "validate_required_folder_files",
                 return_value=self.validation_result,
             ), \
             patch.object(
                 pipeline,
                 "_create_run_folder",
                 side_effect=self._fake_create_run_folder,
             ), \
             patch.object(
                 pipeline,
                 "download_file_from_drive",
                 side_effect=fake_download_file_from_drive,
             ), \
             patch.object(
                 pipeline,
                 "read_acta_data_xlsx",
                 return_value={"numero de acta": "1"},
             ), \
             patch.object(
                 pipeline,
                 "generate_anexo_qrs",
                 side_effect=fail_generating_qrs,
             ):
            with self.assertRaises(RuntimeError) as context:
                pipeline.run_pipeline("https://drive.google.com/folders/test")

        summary = getattr(context.exception, "run_summary", None)
        self.assertIsNotNone(summary)
        self.assertTrue(summary["cleanup"]["attempted"])
        self.assertTrue(summary["cleanup"]["deleted"])
        self.assertFalse(Path(summary["run_folder"]).exists())

    def test_successful_run_deletes_temp_folder_and_keeps_summary(self):
        def fake_download_file_from_drive(service, file_id, destination_path):
            return self._write_file(Path(destination_path), "xlsx")

        def fake_generate_anexo_qrs(validation_result, output_folder):
            output_folder = Path(output_folder)
            return {
                "anexoA": self._write_file(output_folder / "ANEXOA_QR.png"),
                "anexoB": self._write_file(output_folder / "ANEXOB_QR.png"),
                "anexoC": self._write_file(output_folder / "ANEXOC_QR.png"),
            }

        def fake_generate_relacion_anexos_docx(
            qr_files, output_folder, template_docx_path, output_filename
        ):
            return self._write_file(Path(output_folder) / output_filename)

        def fake_convert_docx_to_pdf(docx_path, output_folder, output_filename):
            return self._write_file(Path(output_folder) / output_filename)

        def fake_upload_file_to_folder(service, folder_id, local_file_path, make_public=False):
            return {"link": f"https://drive.test/{Path(local_file_path).name}"}

        def fake_generate_relacion_anexos_qr(relacion_anexos_link, output_folder):
            return self._write_file(Path(output_folder) / "RelacionAnexosQR_QR.png")

        def fake_generate_acta_docx(
            acta_data, relacion_anexos_qr_path, output_folder, template_docx_path, output_filename
        ):
            return self._write_file(Path(output_folder) / output_filename)

        with patch.object(pipeline, "get_drive_service", return_value=object()), \
             patch.object(
                 pipeline,
                 "validate_required_folder_files",
                 return_value=self.validation_result,
             ), \
             patch.object(
                 pipeline,
                 "_create_run_folder",
                 side_effect=self._fake_create_run_folder,
             ), \
             patch.object(
                 pipeline,
                 "download_file_from_drive",
                 side_effect=fake_download_file_from_drive,
             ), \
             patch.object(
                 pipeline,
                 "read_acta_data_xlsx",
                 return_value={"numero de acta": "1"},
             ), \
             patch.object(
                 pipeline,
                 "generate_anexo_qrs",
                 side_effect=fake_generate_anexo_qrs,
             ), \
             patch.object(
                 pipeline,
                 "generate_relacion_anexos_docx",
                 side_effect=fake_generate_relacion_anexos_docx,
             ), \
             patch.object(
                 pipeline,
                 "convert_docx_to_pdf",
                 side_effect=fake_convert_docx_to_pdf,
             ), \
             patch.object(
                 pipeline,
                 "upload_file_to_folder",
                 side_effect=fake_upload_file_to_folder,
             ), \
             patch.object(
                 pipeline,
                 "generate_relacion_anexos_qr",
                 side_effect=fake_generate_relacion_anexos_qr,
             ), \
             patch.object(
                 pipeline,
                 "generate_acta_docx",
                 side_effect=fake_generate_acta_docx,
             ):
            result = pipeline.run_pipeline("https://drive.google.com/folders/test")

        self.assertEqual(result["folder_name"], "Restaurant Norte")
        self.assertTrue(result["cleanup"]["attempted"])
        self.assertTrue(result["cleanup"]["deleted"])
        self.assertFalse(Path(result["run_folder"]).exists())
        self.assertEqual(len(result["generated_files"]), 10)
        self.assertEqual(len(result["uploaded_files"]), 9)

    def test_cleanup_retries_and_reports_warning_when_delete_keeps_failing(self):
        run_folder = Path(tempfile.mkdtemp(dir=self.runs_path))
        self.created_run_folders.append(run_folder)

        with patch.object(pipeline.shutil, "rmtree", side_effect=PermissionError("locked")), \
             patch.object(pipeline.time, "sleep"):
            cleanup = pipeline._delete_run_folder(run_folder, retries=3, delay_seconds=0)

        self.assertTrue(cleanup["attempted"])
        self.assertFalse(cleanup["deleted"])
        self.assertIn("PermissionError", cleanup["error"])
        self.assertTrue(run_folder.exists())


if __name__ == "__main__":
    unittest.main()
