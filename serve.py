"""
serve.py - Lightweight local web server for LTE KPI Dashboard
Runs on standard library Python with no external dependencies.
"""
import http.server
import json
import socketserver
import subprocess
import threading
import time
import webbrowser
import os
import sys
from urllib.parse import urlparse

DATA_UPDATE_LOCK = threading.Lock()


class DashboardHandler(http.server.SimpleHTTPRequestHandler):
    def do_GET(self):
        if self.path == "/api/data-version":
            names = ("userReportDataCompact.js", "daywiseData.js", "kpiDataSources.js")
            versions = {}
            for name in names:
                try:
                    stat = os.stat(name)
                    versions[name] = [stat.st_mtime_ns, stat.st_size]
                except OSError:
                    versions[name] = None
            try:
                with open("kpiDataSources.json", encoding="utf-8") as manifest_file:
                    updated_at = json.load(manifest_file).get("updatedAt")
            except (OSError, ValueError):
                updated_at = None
            versions["updatedAt"] = updated_at
            payload = json.dumps(versions).encode("utf-8")
            self.send_response(200)
            self.send_header("Content-Type", "application/json; charset=utf-8")
            self.send_header("Cache-Control", "no-store")
            self.send_header("Content-Length", str(len(payload)))
            self.end_headers()
            self.wfile.write(payload)
            return
        super().do_GET()

    def do_POST(self):
        if self.path != "/api/sync-data":
            self.send_error(404)
            return
        origin = urlparse(self.headers.get("Origin", ""))
        if not origin.netloc or origin.netloc != self.headers.get("Host") or origin.scheme != "http":
            self._send_json(403, {"message": "Sync requests must come from this dashboard."})
            return

        directory = os.path.dirname(os.path.abspath(__file__))
        updater = os.path.join(directory, "update_kpi_data.py")
        try:
            with DATA_UPDATE_LOCK:
                try:
                    with open(os.path.join(directory, "kpiDataSources.json"), encoding="utf-8") as manifest_file:
                        previous_updated_at = json.load(manifest_file).get("updatedAt")
                except (OSError, ValueError):
                    previous_updated_at = None

                result = subprocess.run(
                    [sys.executable, updater, "--once"],
                    cwd=directory,
                    check=False,
                    capture_output=True,
                    text=True,
                    encoding="utf-8",
                    errors="replace",
                    timeout=600,
                )
                if result.returncode:
                    details = (result.stderr or result.stdout).strip().splitlines()
                    message = details[-1] if details else f"Importer exited with status {result.returncode}."
                    self._send_json(500, {"message": f"Sync failed: {message}"})
                    return

                try:
                    with open(os.path.join(directory, "kpiDataSources.json"), encoding="utf-8") as manifest_file:
                        manifest = json.load(manifest_file)
                except (OSError, ValueError) as error:
                    self._send_json(500, {"message": f"Sync finished, but report metadata could not be read: {error}"})
                    return

            output = result.stdout.lower()
            if "still syncing" in output:
                self._send_json(202, {
                    "pending": True,
                    "message": "OneDrive is still syncing a workbook. Wait a moment, then press SYNC again.",
                })
                return
            changed = manifest.get("updatedAt") != previous_updated_at
            latest_dates = [
                manifest.get("hourlyLatestDate"),
                manifest.get("daywiseLatestDate"),
            ]
            latest_dates = [date for date in latest_dates if date]
            latest_date = max(latest_dates, default=None)
            self._send_json(200, {
                "changed": changed,
                "latestDate": latest_date,
                "message": (
                    f"Synced successfully through {latest_date}."
                    if changed else f"Already up to date through {latest_date}."
                ),
            })
        except subprocess.TimeoutExpired:
            self._send_json(504, {"message": "Sync timed out. Check OneDrive availability and try again."})
        except OSError as error:
            self._send_json(500, {"message": f"Could not start the KPI importer: {error}"})

    def _send_json(self, status_code, data):
        payload = json.dumps(data).encode("utf-8")
        self.send_response(status_code)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Cache-Control", "no-store")
        self.send_header("Content-Length", str(len(payload)))
        self.end_headers()
        self.wfile.write(payload)


def start_data_monitor(directory):
    updater = os.path.join(directory, "update_kpi_data.py")

    def monitor():
        while True:
            try:
                with DATA_UPDATE_LOCK:
                    result = subprocess.run(
                        [sys.executable, updater, "--once"],
                        cwd=directory,
                        check=False,
                    )
                if result.returncode:
                    print(f"[KPI monitor] Importer exited with status {result.returncode}.", flush=True)
            except OSError as error:
                print(f"[KPI monitor] Could not start importer: {error}", flush=True)
            time.sleep(300)

    threading.Thread(target=monitor, name="kpi-data-monitor", daemon=True).start()


def run(port=8080):
    directory = os.path.dirname(os.path.abspath(__file__))
    os.chdir(directory)

    start_data_monitor(directory)
    handler = DashboardHandler
    # Add CORS headers if needed for local fetch
    handler.extensions_map.update({
        '.js': 'application/javascript',
        '.css': 'text/css',
        '.html': 'text/html',
        '.json': 'application/json',
    })

    while port < 8200:
        try:
            with socketserver.ThreadingTCPServer(("127.0.0.1", port), handler) as httpd:
                httpd.daemon_threads = True
                url = f"http://127.0.0.1:{port}/index.html"
                print(f"=====================================================")
                print(f"  LTE KPI Dashboard is running live!")
                print(f"  Local URL: {url}")
                print(f"  Directory: {directory}")
                print(f"=====================================================")
                try:
                    webbrowser.open(url)
                except Exception:
                    pass
                httpd.serve_forever()
        except (OSError, PermissionError) as e:
            port += 1
            if port >= 8200:
                raise e

if __name__ == "__main__":
    initial_port = int(sys.argv[1]) if len(sys.argv) > 1 else 8080
    run(initial_port)
