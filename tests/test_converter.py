from pathlib import Path
import subprocess
import sys
from types import SimpleNamespace

import pytest

from doc_fix.converter import ConversionError, LibreOfficeConverter, WordConverter


def test_normalize_docx_copies_to_output(tmp_path: Path) -> None:
    source = tmp_path / "source.docx"
    source.write_bytes(b"fake docx bytes")

    result = WordConverter().normalize_to_docx(source, tmp_path / "out", "template")

    assert result.converted is False
    assert result.docx_path.name == "template.converted.docx"
    assert result.docx_path.read_bytes() == b"fake docx bytes"


def test_normalize_rejects_unsupported_file_type(tmp_path: Path) -> None:
    source = tmp_path / "source.txt"
    source.write_text("not a document", encoding="utf-8")

    with pytest.raises(ConversionError):
        WordConverter().normalize_to_docx(source, tmp_path / "out", "input")


def test_word_converter_initializes_com_for_worker_threads(tmp_path: Path, monkeypatch) -> None:
    source = tmp_path / "source.doc"
    source.write_bytes(b"doc bytes")
    output = tmp_path / "source.docx"
    events: list[str] = []

    class FakeDocument:
        def SaveAs2(self, path: str, FileFormat: int) -> None:
            events.append(f"save:{FileFormat}")
            Path(path).write_bytes(b"converted")

        def Close(self, save_changes: bool) -> None:
            events.append(f"close:{save_changes}")

    class FakeDocuments:
        def Open(self, path: str, confirm_conversions: bool, read_only: bool):
            events.append(f"open:{read_only}")
            return FakeDocument()

    class FakeWord:
        def __init__(self) -> None:
            self.Documents = FakeDocuments()
            self.Visible = True
            self.DisplayAlerts = 1

        def Quit(self) -> None:
            events.append("quit")

    def co_initialize() -> None:
        events.append("coinit")

    def co_uninitialize() -> None:
        events.append("couninit")

    def dispatch_ex(name: str):
        events.append(f"dispatch:{name}")
        return FakeWord()

    fake_pythoncom = SimpleNamespace(CoInitialize=co_initialize, CoUninitialize=co_uninitialize)
    fake_client = SimpleNamespace(DispatchEx=dispatch_ex)
    fake_win32com = SimpleNamespace(client=fake_client)
    monkeypatch.setitem(sys.modules, "pythoncom", fake_pythoncom)
    monkeypatch.setitem(sys.modules, "win32com", fake_win32com)
    monkeypatch.setitem(sys.modules, "win32com.client", fake_client)

    WordConverter()._convert_with_word(source, output, 16)

    assert output.read_bytes() == b"converted"
    assert events == [
        "coinit",
        "dispatch:Word.Application",
        "open:True",
        "save:16",
        "close:False",
        "quit",
        "couninit",
    ]


def test_libreoffice_normalize_doc_converts_to_docx(tmp_path: Path, monkeypatch) -> None:
    source = tmp_path / "source.doc"
    source.write_bytes(b"doc bytes")

    def fake_run(command, **kwargs):
        out_dir = Path(command[command.index("--outdir") + 1])
        input_path = Path(command[-1])
        (out_dir / f"{input_path.stem}.docx").write_bytes(b"converted docx")
        return subprocess.CompletedProcess(command, 0, stdout="ok", stderr="")

    monkeypatch.setattr(subprocess, "run", fake_run)

    result = LibreOfficeConverter(binary="soffice-test").normalize_to_docx(source, tmp_path / "out", "input")

    assert result.converted is True
    assert result.docx_path == tmp_path / "out" / "input.converted.docx"
    assert result.docx_path.read_bytes() == b"converted docx"


def test_libreoffice_export_to_doc(tmp_path: Path, monkeypatch) -> None:
    source = tmp_path / "input.corrected.docx"
    source.write_bytes(b"docx bytes")

    def fake_run(command, **kwargs):
        out_dir = Path(command[command.index("--outdir") + 1])
        input_path = Path(command[-1])
        (out_dir / f"{input_path.stem}.doc").write_bytes(b"converted doc")
        return subprocess.CompletedProcess(command, 0, stdout="ok", stderr="")

    monkeypatch.setattr(subprocess, "run", fake_run)

    output_path = LibreOfficeConverter().export_to_doc(source, tmp_path / "final.doc")

    assert output_path == tmp_path / "final.doc"
    assert output_path.read_bytes() == b"converted doc"


def test_libreoffice_raises_on_command_failure(tmp_path: Path, monkeypatch) -> None:
    source = tmp_path / "source.doc"
    source.write_bytes(b"doc bytes")

    def fake_run(command, **kwargs):
        return subprocess.CompletedProcess(command, 1, stdout="", stderr="conversion failed")

    monkeypatch.setattr(subprocess, "run", fake_run)

    with pytest.raises(ConversionError, match="conversion failed"):
        LibreOfficeConverter().normalize_to_docx(source, tmp_path / "out", "input")


def test_libreoffice_raises_when_output_is_missing(tmp_path: Path, monkeypatch) -> None:
    source = tmp_path / "source.doc"
    source.write_bytes(b"doc bytes")

    def fake_run(command, **kwargs):
        return subprocess.CompletedProcess(command, 0, stdout="no output", stderr="")

    monkeypatch.setattr(subprocess, "run", fake_run)

    with pytest.raises(ConversionError, match="did not create expected output"):
        LibreOfficeConverter().normalize_to_docx(source, tmp_path / "out", "input")
