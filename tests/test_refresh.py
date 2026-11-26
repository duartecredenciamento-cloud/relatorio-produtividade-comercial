import os
import tempfile
import unittest
import uuid
from pathlib import Path
from unittest.mock import patch

_TEST_DATA = Path(__file__).resolve().parents[1] / f".test_appdata_{uuid.uuid4().hex}"
_TEST_DATA.mkdir(parents=True, exist_ok=True)
os.environ["APPDATA_DIR"] = str(_TEST_DATA)
import app_server


class RefreshFailureTests(unittest.TestCase):
    def test_drive_sync_failure_does_not_leave_refresh_stuck(self):
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp) / "matrizes"
            root.mkdir()
            with patch.dict(os.environ, {
                "MATRICES_DIR": str(root),
                "RESPONSIBLES_FILE": str(Path(tmp) / "missing_roster.xlsx"),
            }), patch.object(app_server, "sync_drive_matrices", side_effect=RuntimeError("falha simulada")):
                store = app_server.DashboardStore()
                store.refresh(force=True)
            self.assertFalse(store.refreshing)
            self.assertIn("Google Drive", store.last_message)
            self.assertIn("falha simulada", store.last_error)


if __name__ == "__main__":
    unittest.main()
