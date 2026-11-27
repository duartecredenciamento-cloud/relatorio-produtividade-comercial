import http.client
import json
import os
import shutil
import tempfile
import threading
import unittest
import uuid
from http.server import ThreadingHTTPServer
from pathlib import Path
from unittest.mock import patch

_TEST_DATA = Path(tempfile.mkdtemp(prefix="duarte-auth-tests-"))
os.environ["APPDATA_DIR"] = str(_TEST_DATA)
import app_server


class AuthFlowTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.admin_user = "admin_" + uuid.uuid4().hex[:10]
        cls.admin_password = uuid.uuid4().hex + uuid.uuid4().hex
        cls.invite_code = uuid.uuid4().hex + uuid.uuid4().hex
        cls.env = patch.dict(os.environ, {
            "DASHBOARD_USER": cls.admin_user,
            "DASHBOARD_PASSWORD": cls.admin_password,
            "DASHBOARD_INVITE_CODE": cls.invite_code,
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
        shutil.rmtree(_TEST_DATA, ignore_errors=True)

    def request(self, method, path, body=None, cookie=None, https=False):
        conn = http.client.HTTPConnection("127.0.0.1", self.port, timeout=5)
        headers = {}
        if body is not None:
            headers["Content-Type"] = "application/json"
            body = json.dumps(body)
        if cookie:
            headers["Cookie"] = cookie
        if https:
            headers["X-Forwarded-Proto"] = "https"
        conn.request(method, path, body=body, headers=headers)
        response = conn.getresponse()
        payload = response.read()
        result = (response.status, dict(response.getheaders()), payload)
        conn.close()
        return result

    def test_exact_routes_and_private_apis(self):
        status, headers, _ = self.request("GET", "/")
        self.assertEqual(status, 303)
        self.assertEqual(headers.get("Location"), "/login.html")
        for path in ("/login.html", "/assets/login.css", "/assets/login.js", "/assets/brand-logo.png"):
            with self.subTest(path=path):
                status, _, body = self.request("GET", path)
                self.assertEqual(status, 200)
                self.assertTrue(body)
        status, headers, _ = self.request("GET", "/painelgeral.html")
        self.assertEqual(status, 303)
        self.assertEqual(headers.get("Location"), "/login.html")
        for method, path, body in (
            ("GET", "/api/dashboard", None),
            ("GET", "/api/session", None),
            ("POST", "/api/refresh", {}),
            ("POST", "/api/marker", {"record_id": "x", "marked": "Sim"}),
            ("POST", "/api/goals", {"service": "CREDENCIAMENTO", "target": 1}),
        ):
            with self.subTest(method=method, path=path):
                status, _, _ = self.request(method, path, body)
                self.assertEqual(status, 401)

    def test_login_invalid_valid_panel_and_server_logout(self):
        status, _, payload = self.request("POST", "/api/login", {"username": self.admin_user, "password": "incorreta"})
        self.assertEqual(status, 401)
        self.assertIn("Usuário ou senha incorretos", payload.decode())
        status, headers, payload = self.request("POST", "/api/login", {"username": self.admin_user, "password": self.admin_password}, https=True)
        self.assertEqual(status, 200)
        self.assertEqual(json.loads(payload)["redirect"], "/painelgeral.html")
        set_cookie = headers["Set-Cookie"]
        self.assertIn("HttpOnly", set_cookie)
        self.assertIn("SameSite=Strict", set_cookie)
        self.assertIn("Secure", set_cookie)
        cookie = set_cookie.split(";", 1)[0]
        status, headers, _ = self.request("GET", "/", cookie=cookie)
        self.assertEqual(status, 303)
        self.assertEqual(headers.get("Location"), "/painelgeral.html")
        status, headers, body = self.request("GET", "/login.html", cookie=cookie)
        self.assertEqual(status, 303)
        self.assertEqual(headers.get("Location"), "/painelgeral.html")
        status, _, body = self.request("GET", "/painelgeral.html", cookie=cookie)
        self.assertEqual(status, 200)
        html = body.decode()
        for path in ("/assets/styles.css", "/assets/app.css", "/assets/data-service.js", "/assets/app.js", "/assets/brand-logo.png"):
            self.assertIn(path, html)
            asset_status, _, asset_body = self.request("GET", path, cookie=cookie)
            self.assertEqual(asset_status, 200)
            self.assertTrue(asset_body)
        status, _, body = self.request("GET", "/api/dashboard", cookie=cookie)
        self.assertEqual(status, 200)
        status, _, _ = self.request("POST", "/api/logout", cookie=cookie)
        self.assertEqual(status, 200)
        status, headers, _ = self.request("GET", "/painelgeral.html", cookie=cookie)
        self.assertEqual(status, 303)
        self.assertEqual(headers.get("Location"), "/login.html")
        self.assertEqual(self.request("GET", "/api/dashboard", cookie=cookie)[0], 401)

    def test_invite_required_and_account_password_is_hashed(self):
        user = "user_" + uuid.uuid4().hex[:10]
        password = uuid.uuid4().hex + uuid.uuid4().hex
        payload = {"username": user, "password": password, "inviteCode": "invalid"}
        with patch.dict(os.environ, {"DASHBOARD_INVITE_CODE": ""}):
            self.assertEqual(self.request("POST", "/api/register", payload)[0], 403)
        self.assertEqual(self.request("POST", "/api/register", payload)[0], 403)
        payload["inviteCode"] = self.invite_code
        status, headers, body = self.request("POST", "/api/register", payload)
        self.assertEqual(status, 201)
        self.assertEqual(json.loads(body)["redirect"], "/painelgeral.html")
        account_text = app_server._ACCOUNTS_FILE.read_text(encoding="utf-8")
        self.assertNotIn(password, account_text)
        self.assertNotIn(self.invite_code, account_text)
        cookie = headers["Set-Cookie"].split(";", 1)[0]
        self.assertEqual(self.request("GET", "/painelgeral.html", cookie=cookie)[0], 200)
        self.request("POST", "/api/logout", cookie=cookie)
        status, headers, _ = self.request("POST", "/api/login", {"username": user, "password": password})
        self.assertEqual(status, 200)
        self.assertEqual(self.request("GET", "/painelgeral.html", cookie=headers["Set-Cookie"].split(";", 1)[0])[0], 200)


if __name__ == "__main__":
    unittest.main()
