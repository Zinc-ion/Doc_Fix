from __future__ import annotations

from http.client import HTTPConnection
from pathlib import Path
import threading

from doc_fix.model import CorrectionReport
from doc_fix.service import CorrectionArtifacts
import doc_fix.web.server as web_server
from doc_fix.web.server import WebSettings, create_server


def test_web_rejects_missing_file(tmp_path: Path) -> None:
    with running_server(tmp_path, fake_correction_runner) as port:
        status, body = post_multipart(port, fields={"note": ("", b"hello")})

    assert status == 400
    assert "No file uploaded" in body


def test_web_rejects_unsupported_extension(tmp_path: Path) -> None:
    with running_server(tmp_path, fake_correction_runner) as port:
        status, body = post_multipart(port, fields={"file": ("bad.txt", b"plain text")})

    assert status == 400
    assert "Only .doc and .docx files are supported" in body


def test_web_upload_success_returns_downloads(tmp_path: Path) -> None:
    with running_server(tmp_path, fake_correction_runner) as port:
        status, body = post_multipart(port, fields={"file": ("target.doc", b"doc bytes")})
        assert status == 200
        assert "下载修正后的 .doc" in body
        doc_link = first_link_for_kind(body, "doc")

        conn = HTTPConnection("127.0.0.1", port, timeout=5)
        conn.request("GET", doc_link)
        response = conn.getresponse()
        data = response.read()
        conn.close()

    assert response.status == 200
    assert data == b"corrected doc"
    assert response.getheader("Content-Disposition") == 'attachment; filename="input.corrected.doc"'


def test_web_download_rejects_path_traversal(tmp_path: Path) -> None:
    with running_server(tmp_path, fake_correction_runner) as port:
        conn = HTTPConnection("127.0.0.1", port, timeout=5)
        conn.request("GET", "/download/../../secret/doc")
        response = conn.getresponse()
        body = response.read().decode("utf-8")
        conn.close()

    assert response.status == 404
    assert "requested file does not exist" in body


def test_web_uses_docx_when_doc_export_failed(tmp_path: Path) -> None:
    with running_server(tmp_path, fake_docx_only_runner) as port:
        status, body = post_multipart(port, fields={"file": ("target.docx", b"docx bytes")})

    assert status == 200
    assert "下载修正后的 .docx" in body
    assert ".doc 导出失败" in body


def test_web_auto_converter_uses_word_on_windows(monkeypatch, tmp_path: Path) -> None:
    captured = {}

    def fake_run_correction(template_path, input_path, out_dir, converter):
        captured["converter"] = converter
        return fake_correction_runner(input_path, out_dir, WebSettings(template_path=template_path))

    monkeypatch.setattr(web_server.platform, "system", lambda: "Windows")
    monkeypatch.setattr(web_server, "run_correction", fake_run_correction)

    web_server.correct_uploaded_document(tmp_path / "input.doc", tmp_path, WebSettings(template_path=tmp_path / "template.doc"))

    assert captured["converter"].__class__.__name__ == "WordConverter"


def test_web_auto_converter_uses_libreoffice_off_windows(monkeypatch, tmp_path: Path) -> None:
    captured = {}

    def fake_run_correction(template_path, input_path, out_dir, converter):
        captured["converter"] = converter
        return fake_correction_runner(input_path, out_dir, WebSettings(template_path=template_path))

    monkeypatch.setattr(web_server.platform, "system", lambda: "Linux")
    monkeypatch.setattr(web_server, "run_correction", fake_run_correction)

    web_server.correct_uploaded_document(tmp_path / "input.doc", tmp_path, WebSettings(template_path=tmp_path / "template.doc"))

    assert captured["converter"].__class__.__name__ == "LibreOfficeConverter"


def test_web_converter_can_be_forced_to_word(monkeypatch, tmp_path: Path) -> None:
    captured = {}

    def fake_run_correction(template_path, input_path, out_dir, converter):
        captured["converter"] = converter
        return fake_correction_runner(input_path, out_dir, WebSettings(template_path=template_path))

    monkeypatch.setattr(web_server.platform, "system", lambda: "Linux")
    monkeypatch.setattr(web_server, "run_correction", fake_run_correction)

    settings = WebSettings(template_path=tmp_path / "template.doc", converter_backend="word")
    web_server.correct_uploaded_document(tmp_path / "input.doc", tmp_path, settings)

    assert captured["converter"].__class__.__name__ == "WordConverter"


class running_server:
    def __init__(self, tmp_path: Path, runner):
        self.settings = WebSettings(
            host="127.0.0.1",
            port=0,
            output_root=tmp_path / "web-output",
            max_upload_bytes=1024 * 1024,
            retention_seconds=24 * 60 * 60,
            template_path=tmp_path / "template.doc",
        )
        self.runner = runner
        self.server = None
        self.thread = None

    def __enter__(self) -> int:
        self.server = create_server(self.settings, correction_runner=self.runner)
        self.thread = threading.Thread(target=self.server.serve_forever, daemon=True)
        self.thread.start()
        return self.server.server_address[1]

    def __exit__(self, exc_type, exc, tb) -> None:
        self.server.shutdown()
        self.server.server_close()
        self.thread.join(timeout=5)


def fake_correction_runner(input_path: Path, out_dir: Path, settings: WebSettings) -> CorrectionArtifacts:
    corrected_docx_path = out_dir / "input.corrected.docx"
    corrected_doc_path = out_dir / "input.corrected.doc"
    correction_report_path = out_dir / "correction_report.json"
    corrected_check_report_path = out_dir / "corrected_check_report.json"
    corrected_doc_check_report_path = out_dir / "corrected_doc_check_report.json"
    corrected_docx_path.write_bytes(b"corrected docx")
    corrected_doc_path.write_bytes(b"corrected doc")
    correction_report_path.write_text("{}", encoding="utf-8")
    corrected_check_report_path.write_text("{}", encoding="utf-8")
    corrected_doc_check_report_path.write_text("{}", encoding="utf-8")
    report = CorrectionReport(
        template_path=str(settings.template_path),
        input_path=str(input_path),
        template_docx_path=str(out_dir / "template.converted.docx"),
        input_docx_path=str(out_dir / "input.converted.docx"),
        corrected_docx_path=str(corrected_docx_path),
        corrected_doc_path=str(corrected_doc_path),
        corrected_check_report_path=str(corrected_check_report_path),
        corrected_doc_check_report_path=str(corrected_doc_check_report_path),
    )
    return CorrectionArtifacts(
        correction_report=report,
        corrected_docx_path=corrected_docx_path,
        corrected_doc_path=corrected_doc_path,
        correction_report_path=correction_report_path,
        corrected_check_report_path=corrected_check_report_path,
        corrected_doc_check_report_path=corrected_doc_check_report_path,
    )


def fake_docx_only_runner(input_path: Path, out_dir: Path, settings: WebSettings) -> CorrectionArtifacts:
    artifacts = fake_correction_runner(input_path, out_dir, settings)
    artifacts.corrected_doc_path.unlink()
    report = CorrectionReport(
        template_path=artifacts.correction_report.template_path,
        input_path=artifacts.correction_report.input_path,
        template_docx_path=artifacts.correction_report.template_docx_path,
        input_docx_path=artifacts.correction_report.input_docx_path,
        corrected_docx_path=artifacts.correction_report.corrected_docx_path,
        corrected_doc_export_error="LibreOffice export failed",
        corrected_check_report_path=artifacts.correction_report.corrected_check_report_path,
    )
    return CorrectionArtifacts(
        correction_report=report,
        corrected_docx_path=artifacts.corrected_docx_path,
        corrected_doc_path=None,
        correction_report_path=artifacts.correction_report_path,
        corrected_check_report_path=artifacts.corrected_check_report_path,
        corrected_doc_check_report_path=None,
    )


def post_multipart(port: int, fields: dict[str, tuple[str, bytes]]) -> tuple[int, str]:
    boundary = "----docfixboundary"
    body = multipart_body(boundary, fields)
    conn = HTTPConnection("127.0.0.1", port, timeout=5)
    conn.request(
        "POST",
        "/correct",
        body=body,
        headers={
            "Content-Type": f"multipart/form-data; boundary={boundary}",
            "Content-Length": str(len(body)),
        },
    )
    response = conn.getresponse()
    data = response.read().decode("utf-8")
    conn.close()
    return response.status, data


def multipart_body(boundary: str, fields: dict[str, tuple[str, bytes]]) -> bytes:
    parts: list[bytes] = []
    for name, (filename, content) in fields.items():
        disposition = f'Content-Disposition: form-data; name="{name}"'
        if filename:
            disposition += f'; filename="{filename}"'
        parts.extend(
            [
                f"--{boundary}\r\n".encode("ascii"),
                f"{disposition}\r\n".encode("utf-8"),
                b"Content-Type: application/octet-stream\r\n\r\n",
                content,
                b"\r\n",
            ]
        )
    parts.append(f"--{boundary}--\r\n".encode("ascii"))
    return b"".join(parts)


def first_link_for_kind(body: str, kind: str) -> str:
    marker = f"/download/"
    index = body.index(marker)
    while index >= 0:
        end = body.index('"', index)
        link = body[index:end]
        if link.endswith(f"/{kind}"):
            return link
        index = body.find(marker, end)
    raise AssertionError(f"No download link for {kind}")
