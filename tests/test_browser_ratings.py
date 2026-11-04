import json
import sys
import threading
import unittest
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from urllib.error import HTTPError

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))
import allods_hpi_to_xlsx as exporter


class RatingsHandler(BaseHTTPRequestHandler):
    def do_GET(self):
        if self.path == "/denied":
            self.send_body(403, b"Access denied")
        elif self.path == "/invalid":
            self.send_body(200, b"<html>Service unavailable</html>")
        elif self.path == "/stuck":
            self.send_body(403, b"<script>window.bp_chl_vars = {};</script>")
        elif "ratings_session=ready" not in self.headers.get("Cookie", ""):
            self.server.challenges += 1
            self.send_body(403, b"""<script>
                window.bp_chl_vars = {};
                document.cookie = 'ratings_session=ready; path=/';
                location.reload();
            </script>""")
        else:
            data = [{"name": "Player", "guild": "A &amp; B <C>", "rating_index": 1, "achievement": 123,
                     "difficulty": 740033086, "time": 456, "shard_id": 601}]
            # The official API also serves JSON with a text/html content type.
            self.send_body(200, json.dumps(data).encode("utf-8"))

    def send_body(self, status, body):
        self.send_response(status)
        self.send_header("Content-Type", "text/html;charset=UTF-8")
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)

    def log_message(self, *args):
        pass


class BrowserRatingsTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.server = ThreadingHTTPServer(("127.0.0.1", 0), RatingsHandler)
        cls.server.challenges = 0
        cls.thread = threading.Thread(target=cls.server.serve_forever, daemon=True)
        cls.thread.start()
        cls.base_url = f"http://127.0.0.1:{cls.server.server_port}"

    @classmethod
    def tearDownClass(cls):
        cls.server.shutdown()
        cls.server.server_close()
        cls.thread.join()

    def client(self):
        self.assertTrue(
            hasattr(exporter, "BrowserRatingsClient"),
            "Ratings loader must execute the site's JavaScript in a browser session",
        )
        return exporter.BrowserRatingsClient()

    def test_executes_javascript_and_preserves_json_fields(self):
        with self.client() as client:
            data = client.fetch_json(self.base_url + "/ratings", 5)
        self.assertEqual(data[0]["difficulty"], 740033086)
        self.assertEqual(data[0]["time"], 456)
        self.assertEqual(data[0]["achievement"], 123)

    def test_reuses_session_for_subsequent_requests(self):
        initial = self.server.challenges
        with self.client() as client:
            client.fetch_json(self.base_url + "/first", 5)
            client.fetch_json(self.base_url + "/second", 5)
        self.assertEqual(self.server.challenges - initial, 1)

    def test_preserves_json_strings_when_content_type_is_html(self):
        with self.client() as client:
            data = client.fetch_json(self.base_url + "/ratings", 5)
        self.assertEqual(data[0]["guild"], "A &amp; B <C>")

    def test_plain_http_denial_is_not_reported_as_json(self):
        with self.client() as client:
            with self.assertRaises(HTTPError) as error:
                client.fetch_json(self.base_url + "/denied", 5)
        self.assertEqual(error.exception.code, 403)

    def test_html_success_is_rejected(self):
        with self.client() as client:
            with self.assertRaisesRegex(RuntimeError, "JSON"):
                client.fetch_json(self.base_url + "/invalid", 5)

    def test_stuck_security_check_has_bounded_timeout(self):
        with self.client() as client:
            with self.assertRaisesRegex(RuntimeError, "security check"):
                client.fetch_json(self.base_url + "/stuck", 1)

    def test_invalid_timeout_is_rejected(self):
        with self.client() as client:
            with self.assertRaises(ValueError):
                client.fetch_json(self.base_url + "/ratings", 0)

    def test_closes_session_after_exception(self):
        with self.assertRaisesRegex(RuntimeError, "report failed"):
            with self.client() as client:
                client.fetch_json(self.base_url + "/ratings", 5)
                raise RuntimeError("report failed")
        with self.assertRaisesRegex(RuntimeError, "context manager"):
            client.fetch_json(self.base_url + "/ratings", 5)


if __name__ == "__main__":
    unittest.main()
