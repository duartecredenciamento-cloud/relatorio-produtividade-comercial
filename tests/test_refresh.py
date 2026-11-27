import os
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

_TEST_DATA = tempfile.TemporaryDirectory(prefix="duarte-refresh-tests-")
os.environ["APPDATA_DIR"] = _TEST_DATA.name
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
            self.assertIn("sincronização indisponível", store.last_error)
            self.assertNotIn("falha simulada", store.last_error)


if __name__ == "__main__":
    unittest.main()
