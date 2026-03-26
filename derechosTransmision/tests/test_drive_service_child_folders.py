import sys
import types
import unittest
from pathlib import Path


SRC_DIR = Path(__file__).resolve().parents[1] / "src"
if str(SRC_DIR) not in sys.path:
    sys.path.insert(0, str(SRC_DIR))

for module_name in [
    "google",
    "google.auth",
    "google.auth.transport",
    "google.auth.transport.requests",
    "google.oauth2",
    "google.oauth2.credentials",
    "google_auth_oauthlib",
    "google_auth_oauthlib.flow",
    "googleapiclient",
    "googleapiclient.discovery",
    "googleapiclient.http",
]:
    sys.modules.setdefault(module_name, types.ModuleType(module_name))

sys.modules["googleapiclient"].__path__ = []
sys.modules.setdefault("googleapiclient.errors", types.ModuleType("googleapiclient.errors"))

sys.modules["google.auth.transport.requests"].Request = object
sys.modules["google.oauth2.credentials"].Credentials = object
sys.modules["google_auth_oauthlib.flow"].InstalledAppFlow = object
sys.modules["googleapiclient.discovery"].build = lambda *args, **kwargs: None
sys.modules["googleapiclient.http"].MediaFileUpload = object
sys.modules["googleapiclient.http"].MediaIoBaseDownload = object

import drive_service  # noqa: E402


class _FakeFilesResource:
    def __init__(self, responses):
        self.responses = responses
        self.calls = []

    def list(self, **kwargs):
        self.calls.append(kwargs)
        response = self.responses[len(self.calls) - 1]
        return types.SimpleNamespace(execute=lambda: response)


class _FakeDriveService:
    def __init__(self, responses):
        self._files_resource = _FakeFilesResource(responses)

    def files(self):
        return self._files_resource


class DriveServiceChildFoldersTests(unittest.TestCase):
    def test_list_child_folders_collects_all_pages_and_builds_links(self):
        service = _FakeDriveService(
            [
                {
                    "files": [{"id": "one", "name": "Uno"}],
                    "nextPageToken": "page-2",
                },
                {
                    "files": [{"id": "two", "name": "Dos"}],
                },
            ]
        )

        child_folders = drive_service.list_child_folders(
            service,
            "https://drive.google.com/drive/folders/parent-123",
        )

        self.assertEqual(
            child_folders,
            [
                {
                    "id": "one",
                    "name": "Uno",
                    "link": "https://drive.google.com/drive/folders/one",
                },
                {
                    "id": "two",
                    "name": "Dos",
                    "link": "https://drive.google.com/drive/folders/two",
                },
            ],
        )
        self.assertEqual(service._files_resource.calls[0]["pageToken"], None)
        self.assertEqual(service._files_resource.calls[1]["pageToken"], "page-2")


if __name__ == "__main__":
    unittest.main()
