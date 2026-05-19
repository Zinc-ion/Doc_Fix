"""LibreOffice-based document conversion for Linux deployments."""

from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
import shutil
import subprocess

from doc_fix.converter.word import ConversionError, ConversionResult


@dataclass(frozen=True)
class LibreOfficeConverter:
    """Normalize .doc/.docx inputs with LibreOffice headless."""

    binary: str = "soffice"

    def normalize_to_docx(self, input_path: Path, out_dir: Path, label: str) -> ConversionResult:
        """Convert .doc inputs to .docx and copy .docx inputs to a work file."""

        input_path = input_path.resolve()
        out_dir.mkdir(parents=True, exist_ok=True)

        if not input_path.exists():
            raise ConversionError(f"File does not exist: {input_path}")

        suffix = input_path.suffix.lower()
        output_path = out_dir / f"{label}.converted.docx"

        if suffix == ".docx":
            if input_path.resolve() != output_path.resolve():
                shutil.copy2(input_path, output_path)
            return ConversionResult(input_path, output_path, converted=False)

        if suffix != ".doc":
            raise ConversionError(f"Unsupported file type '{input_path.suffix}'. Only .doc and .docx are supported.")

        working_source = out_dir / f"{label}.source.doc"
        shutil.copy2(input_path, working_source)
        converted_path = self._convert_with_libreoffice(working_source, out_dir, "docx", label)
        converted_path.replace(output_path)
        return ConversionResult(input_path, output_path, converted=True)

    def export_to_doc(self, docx_path: Path, output_path: Path) -> Path:
        """Export a .docx file to .doc with LibreOffice headless."""

        docx_path = docx_path.resolve()
        output_path = output_path.resolve()
        output_path.parent.mkdir(parents=True, exist_ok=True)
        converted_path = self._convert_with_libreoffice(docx_path, output_path.parent, "doc", output_path.stem)
        if converted_path.resolve() != output_path:
            converted_path.replace(output_path)
        return output_path

    def _convert_with_libreoffice(
        self,
        input_path: Path,
        out_dir: Path,
        target_extension: str,
        label: str,
    ) -> Path:
        profile_dir = out_dir / f"{label}.lo-profile"
        profile_dir.mkdir(parents=True, exist_ok=True)
        expected_path = out_dir / f"{input_path.stem}.{target_extension}"
        if expected_path.exists():
            expected_path.unlink()

        command = [
            self.binary,
            "--headless",
            f"-env:UserInstallation={profile_dir.resolve().as_uri()}",
            "--convert-to",
            target_extension,
            "--outdir",
            str(out_dir),
            str(input_path),
        ]
        try:
            result = subprocess.run(
                command,
                capture_output=True,
                check=False,
                text=True,
                timeout=120,
            )
        except (OSError, subprocess.TimeoutExpired) as exc:
            raise ConversionError(f"LibreOffice failed to convert '{input_path}': {exc}") from exc

        if result.returncode != 0:
            detail = (result.stderr or result.stdout or "").strip()
            if detail:
                raise ConversionError(f"LibreOffice failed to convert '{input_path}': {detail}")
            raise ConversionError(f"LibreOffice failed to convert '{input_path}' with exit code {result.returncode}.")

        if not expected_path.exists():
            detail = (result.stderr or result.stdout or "").strip()
            suffix = f": {detail}" if detail else ""
            raise ConversionError(f"LibreOffice did not create expected output '{expected_path}'{suffix}")

        return expected_path
