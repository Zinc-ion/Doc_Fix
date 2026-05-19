"""Temporary stdlib web server for fixed-template correction."""

from __future__ import annotations

from dataclasses import dataclass
from html import escape
from http import HTTPStatus
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
import mimetypes
import os
import platform
from pathlib import Path
import re
import shutil
import time
from typing import Callable
from urllib.parse import unquote, urlparse
from uuid import uuid4

from doc_fix.converter import LibreOfficeConverter, WordConverter
from doc_fix.service import CorrectionArtifacts, run_correction


REPO_ROOT = Path(__file__).resolve().parents[2]
FIXED_TEMPLATE_PATH = REPO_ROOT / "example" / "项目申报-模板.doc"
DEFAULT_OUTPUT_ROOT = REPO_ROOT / "output" / "web"
ALLOWED_SUFFIXES = {".doc", ".docx"}
JOB_ID_PATTERN = re.compile(r"^[0-9a-f]{32}$")


@dataclass(frozen=True)
class WebSettings:
    """Runtime settings for the temporary correction server."""

    host: str = "0.0.0.0"
    port: int = 8000
    max_upload_bytes: int = 100 * 1024 * 1024
    retention_seconds: int = 24 * 60 * 60
    output_root: Path = DEFAULT_OUTPUT_ROOT
    libreoffice_binary: str = "soffice"
    converter_backend: str = "auto"
    template_path: Path = FIXED_TEMPLATE_PATH


@dataclass(frozen=True)
class UploadedFile:
    """One uploaded file parsed from a multipart request."""

    filename: str
    data: bytes


CorrectionRunner = Callable[[Path, Path, WebSettings], CorrectionArtifacts]


def load_settings() -> WebSettings:
    """Load web settings from environment variables."""

    max_mb = int(os.environ.get("DOC_FIX_WEB_MAX_MB", "100"))
    retention_hours = int(os.environ.get("DOC_FIX_WEB_RETENTION_HOURS", "24"))
    return WebSettings(
        host=os.environ.get("DOC_FIX_WEB_HOST", "0.0.0.0"),
        port=int(os.environ.get("DOC_FIX_WEB_PORT", "8000")),
        max_upload_bytes=max_mb * 1024 * 1024,
        retention_seconds=retention_hours * 60 * 60,
        libreoffice_binary=os.environ.get("DOC_FIX_LIBREOFFICE_BIN", "soffice"),
        converter_backend=os.environ.get("DOC_FIX_WEB_CONVERTER", "auto"),
    )


def correct_uploaded_document(input_path: Path, out_dir: Path, settings: WebSettings) -> CorrectionArtifacts:
    """Run the fixed-template correction workflow for one uploaded document."""

    converter = _select_converter(settings)
    return run_correction(settings.template_path, input_path, out_dir, converter)


def _select_converter(settings: WebSettings):
    backend = settings.converter_backend.lower()
    if backend == "auto":
        backend = "word" if platform.system() == "Windows" else "libreoffice"
    if backend == "word":
        return WordConverter()
    if backend == "libreoffice":
        return LibreOfficeConverter(binary=settings.libreoffice_binary)
    raise ValueError("DOC_FIX_WEB_CONVERTER must be 'auto', 'word', or 'libreoffice'.")


def build_handler(
    settings: WebSettings,
    correction_runner: CorrectionRunner = correct_uploaded_document,
) -> type[BaseHTTPRequestHandler]:
    """Build a request handler bound to settings and a correction runner."""

    class CorrectionRequestHandler(BaseHTTPRequestHandler):
        server_version = "DocFixWeb/0.1"

        def do_GET(self) -> None:
            path = urlparse(self.path).path
            if path == "/":
                self._send_html(HTTPStatus.OK, _index_page(settings))
                return
            if path.startswith("/download/"):
                self._handle_download(path)
                return
            self._send_html(HTTPStatus.NOT_FOUND, _message_page("Not found", "The requested page does not exist."))

        def do_POST(self) -> None:
            path = urlparse(self.path).path
            if path != "/correct":
                self._send_html(HTTPStatus.NOT_FOUND, _message_page("Not found", "The requested page does not exist."))
                return
            self._handle_correct()

        def log_message(self, format: str, *args: object) -> None:
            return

        def _handle_correct(self) -> None:
            _cleanup_old_jobs(settings.output_root, settings.retention_seconds)
            content_length = int(self.headers.get("Content-Length", "0") or "0")
            if content_length <= 0:
                self._send_html(HTTPStatus.BAD_REQUEST, _message_page("Upload failed", "No upload body was received."))
                return
            if content_length > settings.max_upload_bytes:
                self._send_html(HTTPStatus.REQUEST_ENTITY_TOO_LARGE, _message_page("Upload failed", "The uploaded file is too large."))
                return

            uploaded_file = self._uploaded_file(content_length)
            if uploaded_file is None or not uploaded_file.filename:
                self._send_html(HTTPStatus.BAD_REQUEST, _message_page("Upload failed", "No file uploaded."))
                return

            suffix = Path(uploaded_file.filename).suffix.lower()
            if suffix not in ALLOWED_SUFFIXES:
                self._send_html(HTTPStatus.BAD_REQUEST, _message_page("Upload failed", "Only .doc and .docx files are supported."))
                return

            job_id = uuid4().hex
            job_dir = settings.output_root / job_id
            job_dir.mkdir(parents=True, exist_ok=True)
            upload_path = job_dir / f"input.uploaded{suffix}"
            upload_path.write_bytes(uploaded_file.data)

            try:
                artifacts = correction_runner(upload_path, job_dir, settings)
            except Exception as exc:
                self._send_html(
                    HTTPStatus.INTERNAL_SERVER_ERROR,
                    _message_page("Correction failed", f"The document could not be corrected: {escape(str(exc))}"),
                )
                return

            self._send_html(HTTPStatus.OK, _success_page(job_id, artifacts))

        def _uploaded_file(self, content_length: int) -> UploadedFile | None:
            content_type = self.headers.get("Content-Type", "")
            match = re.search(r"boundary=(?P<boundary>[^;]+)", content_type)
            if match is None:
                return None
            boundary = match.group("boundary").strip('"').encode("utf-8")
            body = self.rfile.read(content_length)
            return _parse_uploaded_file(body, boundary)

        def _handle_download(self, path: str) -> None:
            parts = [unquote(part) for part in path.strip("/").split("/")]
            if len(parts) != 3 or parts[0] != "download":
                self._send_html(HTTPStatus.NOT_FOUND, _message_page("Not found", "The requested file does not exist."))
                return

            job_id, kind = parts[1], parts[2]
            if not JOB_ID_PATTERN.match(job_id):
                self._send_html(HTTPStatus.NOT_FOUND, _message_page("Not found", "The requested file does not exist."))
                return

            file_path = _download_path(settings.output_root, job_id, kind)
            if file_path is None or not file_path.exists():
                self._send_html(HTTPStatus.NOT_FOUND, _message_page("Not found", "The requested file does not exist."))
                return

            self.send_response(HTTPStatus.OK)
            content_type = mimetypes.guess_type(file_path.name)[0] or "application/octet-stream"
            self.send_header("Content-Type", content_type)
            self.send_header("Content-Length", str(file_path.stat().st_size))
            self.send_header("Content-Disposition", f'attachment; filename="{file_path.name}"')
            self.end_headers()
            with file_path.open("rb") as source:
                shutil.copyfileobj(source, self.wfile)

        def _send_html(self, status: HTTPStatus, body: str) -> None:
            data = body.encode("utf-8")
            self.send_response(status)
            self.send_header("Content-Type", "text/html; charset=utf-8")
            self.send_header("Content-Length", str(len(data)))
            self.end_headers()
            self.wfile.write(data)

    return CorrectionRequestHandler


def create_server(
    settings: WebSettings | None = None,
    correction_runner: CorrectionRunner = correct_uploaded_document,
) -> ThreadingHTTPServer:
    """Create a configured HTTP server instance."""

    settings = settings or load_settings()
    settings.output_root.mkdir(parents=True, exist_ok=True)
    handler = build_handler(settings, correction_runner=correction_runner)
    return ThreadingHTTPServer((settings.host, settings.port), handler)


def _cleanup_old_jobs(output_root: Path, retention_seconds: int) -> None:
    if retention_seconds <= 0 or not output_root.exists():
        return
    cutoff = time.time() - retention_seconds
    for child in output_root.iterdir():
        if not child.is_dir():
            continue
        try:
            if child.stat().st_mtime < cutoff:
                shutil.rmtree(child)
        except OSError:
            continue


def _download_path(output_root: Path, job_id: str, kind: str) -> Path | None:
    filenames = {
        "doc": "input.corrected.doc",
        "docx": "input.corrected.docx",
        "report": "correction_report.json",
        "check": "corrected_check_report.json",
        "doc-check": "corrected_doc_check_report.json",
    }
    filename = filenames.get(kind)
    if filename is None:
        return None
    return output_root / job_id / filename


def _parse_uploaded_file(body: bytes, boundary: bytes) -> UploadedFile | None:
    delimiter = b"--" + boundary
    for raw_part in body.split(delimiter):
        part = raw_part
        if not part or part in {b"--", b"--\r\n"}:
            continue
        if part.startswith(b"\r\n"):
            part = part[2:]
        if part.endswith(b"--\r\n"):
            part = part[:-4]
        elif part.endswith(b"--"):
            part = part[:-2]
        if part.endswith(b"\r\n"):
            part = part[:-2]

        header_bytes, separator, data = part.partition(b"\r\n\r\n")
        if not separator:
            continue
        headers = header_bytes.decode("utf-8", errors="replace").split("\r\n")
        disposition = next((line for line in headers if line.lower().startswith("content-disposition:")), "")
        if 'name="file"' not in disposition:
            continue
        filename_match = re.search(r'filename="(?P<filename>[^"]*)"', disposition)
        filename = filename_match.group("filename") if filename_match else ""
        return UploadedFile(filename=filename, data=data)
    return None


def _index_page(settings: WebSettings) -> str:
    max_mb = settings.max_upload_bytes // (1024 * 1024)
    return f"""<!doctype html>
<html lang="zh-CN">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>Doc Fix</title>
  <style>{_css()}</style>
</head>
<body>
  <main>
    <h1>申报书格式修改</h1>
    <form action="/correct" method="post" enctype="multipart/form-data">
      <label class="drop">
        <span>选择 .doc 或 .docx 文件</span>
        <input type="file" name="file" accept=".doc,.docx" required>
      </label>
      <button type="submit">开始修改</button>
    </form>
    <p class="note">固定模板：项目申报-模板.doc。最大上传 {max_mb} MB。</p>
  </main>
</body>
</html>"""


def _success_page(job_id: str, artifacts: CorrectionArtifacts) -> str:
    report = artifacts.correction_report
    preferred_kind = "doc" if artifacts.corrected_doc_path else "docx"
    preferred_label = "下载修正后的 .doc" if artifacts.corrected_doc_path else "下载修正后的 .docx"
    warning = ""
    if report.corrected_doc_export_error:
        warning = f"<p class=\"warning\">.doc 导出失败，已提供 .docx 下载：{escape(report.corrected_doc_export_error)}</p>"
    actions = "".join(
        [
            f'<a class="button" href="/download/{job_id}/{preferred_kind}">{preferred_label}</a>',
            f'<a href="/download/{job_id}/docx">下载 .docx</a>',
            f'<a href="/download/{job_id}/report">下载修正报告 JSON</a>',
            f'<a href="/download/{job_id}/check">下载修正版检查 JSON</a>',
        ]
    )
    if artifacts.corrected_doc_check_report_path:
        actions += f'<a href="/download/{job_id}/doc-check">下载 .doc 导出检查 JSON</a>'
    return f"""<!doctype html>
<html lang="zh-CN">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>修改完成</title>
  <style>{_css()}</style>
</head>
<body>
  <main>
    <h1>修改完成</h1>
    {warning}
    <div class="actions">{actions}</div>
    <a class="back" href="/">继续处理下一个文件</a>
  </main>
</body>
</html>"""


def _message_page(title: str, message: str) -> str:
    return f"""<!doctype html>
<html lang="zh-CN">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>{escape(title)}</title>
  <style>{_css()}</style>
</head>
<body>
  <main>
    <h1>{escape(title)}</h1>
    <p class="warning">{message}</p>
    <a class="back" href="/">返回上传</a>
  </main>
</body>
</html>"""


def _css() -> str:
    return """
* { box-sizing: border-box; }
body {
  margin: 0;
  min-height: 100vh;
  display: grid;
  place-items: center;
  background: #f5f7fb;
  color: #172033;
  font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", "Microsoft YaHei", sans-serif;
}
main {
  width: min(560px, calc(100vw - 32px));
  background: #fff;
  border: 1px solid #d9e0ea;
  border-radius: 8px;
  padding: 32px;
  box-shadow: 0 16px 40px rgba(23, 32, 51, .08);
}
h1 { margin: 0 0 24px; font-size: 28px; line-height: 1.25; }
form, .actions { display: grid; gap: 16px; }
.drop {
  display: grid;
  gap: 12px;
  padding: 24px;
  border: 1px dashed #93a3b8;
  border-radius: 8px;
  background: #fbfcff;
}
input { width: 100%; }
button, .button {
  display: inline-flex;
  align-items: center;
  justify-content: center;
  min-height: 44px;
  border: 0;
  border-radius: 6px;
  background: #1769e0;
  color: #fff;
  font-size: 16px;
  text-decoration: none;
  cursor: pointer;
}
a { color: #1769e0; }
.actions a:not(.button), .back {
  min-height: 36px;
  display: inline-flex;
  align-items: center;
}
.note { color: #5f6f86; margin: 16px 0 0; }
.warning { color: #9a3412; line-height: 1.6; }
"""


def main() -> None:
    """Run the temporary correction web server."""

    settings = load_settings()
    server = create_server(settings)
    print(f"Doc_Fix web server listening on http://{settings.host}:{settings.port}")
    print(f"Fixed template: {settings.template_path}")
    server.serve_forever()


if __name__ == "__main__":
    main()
