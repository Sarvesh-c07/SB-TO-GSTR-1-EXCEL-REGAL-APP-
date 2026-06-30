from __future__ import annotations

import json
import mimetypes
import os
import threading
import uuid
import webbrowser
from email.parser import BytesParser
from email.policy import default
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from urllib.parse import urlparse

from extractor import extract_pdf
from workbook_service import append_rows, create_new_master, install_master, master_status


ROOT = Path(__file__).resolve().parent
STATIC = ROOT / "static"
DATA = ROOT / "data"
PENDING = DATA / "pending"
MASTER = DATA / "master.xlsx"
AUDIT = DATA / "audit.jsonl"
BACKUPS = DATA / "backups"
MAX_UPLOAD = 80 * 1024 * 1024
LOCK = threading.RLock()


def _multipart(headers, body: bytes):
    raw = b"Content-Type: " + headers["Content-Type"].encode() + b"\r\nMIME-Version: 1.0\r\n\r\n" + body
    message = BytesParser(policy=default).parsebytes(raw)
    fields: dict[str, list[dict]] = {}
    for part in message.iter_parts():
        name = part.get_param("name", header="content-disposition")
        if not name:
            continue
        fields.setdefault(name, []).append({
            "filename": part.get_filename(),
            "content": part.get_payload(decode=True) or b"",
        })
    return fields


class Handler(BaseHTTPRequestHandler):
    server_version = "RegalExtract/1.0"

    def log_message(self, fmt, *args):
        print(f"[{self.log_date_time_string()}] {fmt % args}")

    def _json(self, payload, status=200):
        data = json.dumps(payload, ensure_ascii=False).encode("utf-8")
        self.send_response(status)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(data)))
        self.send_header("Cache-Control", "no-store")
        self.end_headers()
        self.wfile.write(data)

    def _error(self, message, status=400):
        self._json({"ok": False, "error": str(message)}, status)

    def _body(self):
        length = int(self.headers.get("Content-Length", "0"))
        if length > MAX_UPLOAD:
            raise ValueError("Upload is too large. The current limit is 80 MB per request.")
        return self.rfile.read(length)

    def do_GET(self):
        path = urlparse(self.path).path
        if path == "/api/status":
            with LOCK:
                return self._json({"ok": True, "master": master_status(MASTER)})
        if path == "/api/download/master":
            if not MASTER.exists():
                return self._error("No master workbook is configured yet.", 404)
            data = MASTER.read_bytes()
            self.send_response(200)
            self.send_header("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            self.send_header("Content-Disposition", 'attachment; filename="3S_PHARMACEUTICALS_MASTER.xlsx"')
            self.send_header("Content-Length", str(len(data)))
            self.end_headers()
            return self.wfile.write(data)
        rel = "index.html" if path == "/" else path.lstrip("/")
        target = (STATIC / rel).resolve()
        if STATIC.resolve() not in target.parents and target != STATIC.resolve():
            return self._error("Not found", 404)
        if not target.exists() or not target.is_file():
            return self._error("Not found", 404)
        data = target.read_bytes()
        mime = mimetypes.guess_type(target.name)[0] or "application/octet-stream"
        self.send_response(200)
        self.send_header("Content-Type", mime)
        self.send_header("Content-Length", str(len(data)))
        self.end_headers()
        self.wfile.write(data)

    def do_POST(self):
        path = urlparse(self.path).path
        try:
            if path == "/api/master":
                fields = _multipart(self.headers, self._body())
                action = (fields.get("action") or [{"content": b""}])[0]["content"].decode().strip()
                with LOCK:
                    DATA.mkdir(parents=True, exist_ok=True)
                    if action == "new":
                        create_new_master(MASTER, BACKUPS)
                    elif action == "existing":
                        uploads = fields.get("master") or []
                        if not uploads or not uploads[0]["content"]:
                            raise ValueError("Choose an Excel master file first.")
                        install_master(uploads[0]["content"], MASTER, BACKUPS)
                    else:
                        raise ValueError("Unknown master action.")
                    status = master_status(MASTER)
                return self._json({"ok": True, "master": status})

            if path == "/api/extract":
                fields = _multipart(self.headers, self._body())
                pdfs = fields.get("pdfs") or []
                if not pdfs:
                    raise ValueError("Choose at least one PDF.")
                rows = []
                for upload in pdfs:
                    if not (upload["filename"] or "").lower().endswith(".pdf"):
                        continue
                    rows.append(extract_pdf(upload["content"], upload["filename"]))
                if not rows:
                    raise ValueError("No PDF files were found in the upload.")
                batch_id = uuid.uuid4().hex
                PENDING.mkdir(parents=True, exist_ok=True)
                (PENDING / f"{batch_id}.json").write_text(json.dumps(rows, ensure_ascii=False), encoding="utf-8")
                return self._json({"ok": True, "batch_id": batch_id, "rows": rows})

            if path == "/api/append":
                payload = json.loads(self._body().decode("utf-8"))
                rows = payload.get("rows") or []
                if not rows:
                    raise ValueError("There are no reviewed rows to append.")
                with LOCK:
                    if not MASTER.exists():
                        raise ValueError("Choose or create a master workbook first.")
                    result = append_rows(MASTER, rows, AUDIT)
                    status = master_status(MASTER)
                return self._json({"ok": True, "result": result, "master": status, "download": "/api/download/master"})
            return self._error("Not found", 404)
        except Exception as exc:
            return self._error(exc, 400)


def main():
    DATA.mkdir(parents=True, exist_ok=True)
    host = os.environ.get("REGAL_HOST", "127.0.0.1")
    port = int(os.environ.get("REGAL_PORT", "8765"))
    server = ThreadingHTTPServer((host, port), Handler)
    url = f"http://{host}:{port}"
    print(f"Regal Extract is running at {url}")
    if os.environ.get("REGAL_NO_BROWSER") != "1":
        threading.Timer(0.7, lambda: webbrowser.open(url)).start()
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        print("\nStopping Regal Extract.")
    finally:
        server.server_close()


if __name__ == "__main__":
    main()
