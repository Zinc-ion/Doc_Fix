from pathlib import Path

from doc_fix.cli.main import _resolve_report_paths


def test_resolve_report_paths_defaults_to_out_dir() -> None:
    paths = _resolve_report_paths(Path("output1"), None, None, None)

    assert paths == (
        Path("output1") / "report.json",
        Path("output1") / "report.md",
        Path("output1") / "report.html",
    )


def test_resolve_report_paths_keeps_explicit_overrides() -> None:
    paths = _resolve_report_paths(
        Path("output1"),
        Path("custom.json"),
        Path("custom.md"),
        Path("custom.html"),
    )

    assert paths == (Path("custom.json"), Path("custom.md"), Path("custom.html"))
