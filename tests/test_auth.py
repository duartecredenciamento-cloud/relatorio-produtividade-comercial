import http.client
import json
import os
import threading
import unittest
import uuid
from pathlib import Path
from http.server import ThreadingHTTPServer
from unittest.mock import patch

_TEST_DATA = Path(__file__).resolve().parents[1] / f".test_appdata_{uuid.uuid4().hex}"
_TEST_DATA.mkdir(parents=True, exist_ok=True)
os.environ["APPDATA_DIR"] = str(_TEST_DATA)
import app_server


class AuthFlowTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.env = patch.dict(os.environ, {
            "DASHBOARD_USER": "tester",
            "DASHBOARD_PASSWORD": "test-password-unique",
            "DUARTE_NO_BROWSER": "1",
        })
        cls.env.start()
        cls.server = ThreadingHTTPServer(("127.0.0.1", 0), app_server.Handler)
        cls.thread = threading.Thread(target=cls.server.serve_forever, daemon=True)
        cls.thread.start()
        cls.port = cls.server.server_address[1]

    @classmethod
    def tearDownClass(cls):
        cls.server.shutdown()
        cls.server.server_close()
        cls.thread.join(timeout=2)
        cls.env.stop()

    def request(self, method, path, body=None, cookie=None):
        conn = http.client.HTTPConnection("127.0.0.1", self.port, timeout=5)
        headers = {}
        if body is not None:
            headers["Content-Type"] = "application/json"
            body = json.dumps(body)
        if cookie:
            headers["Cookie"] = cookie
        conn.request(method, path, body=body, headers=headers)
        response = conn.getresponse()
        payload = response.read()
        result = (response.status, dict(response.getheaders()), payload)
        conn.close()
        return result

    def test_custom_login_protects_dashboard_and_logout_invalidates_cookie(self):
        status, _, _ = self.request("GET", "/")
        self.assertEqual(status, 303)
        status, _, _ = self.request("POST", "/api/login", {"username": "tester", "password": "wrong"})
        self.assertEqual(status, 401)
        status, headers, _ = self.request("POST", "/api/login", {"username": "tester", "password": "test-password-unique"})
        self.assertEqual(status, 200)
        cookie = headers["Set-Cookie"].split(";", 1)[0]
        status, _, _ = self.request("GET", "/api/session", cookie=cookie)
        self.assertEqual(status, 200)
        status, _, _ = self.request("GET", "/", cookie=cookie)
        self.assertEqual(status, 200)
        status, _, _ = self.request("POST", "/api/logout", cookie=cookie)
        self.assertEqual(status, 200)
        status, _, _ = self.request("GET", "/api/session", cookie=cookie)
        self.assertEqual(status, 401)

    def test_login_screen_and_assets_are_public(self):
        for path in ("/login", "/assets/login.css", "/assets/login.js"):
            with self.subTest(path=path):
                status, _, _ = self.request("GET", path)
                self.assertEqual(status, 200)


if __name__ == "__main__":
    unittest.main()
