import subprocess
import tempfile
from pathlib import Path

from .base import BaseConverter
from .docx_converter import DocxConverter


class DocConverter(BaseConverter):
    """Convert DOC → DOCX (LibreOffice) → Markdown (pandoc)."""

    def convert(self, input_path: str, output_path: str, images_dir: str) -> None:
        with tempfile.TemporaryDirectory() as tmpdir:
            docx_path = self._to_docx(input_path, tmpdir)
            DocxConverter().convert(docx_path, output_path, images_dir)

    def _to_docx(self, input_path: str, outdir: str) -> str:
        try:
            subprocess.run(
                [
                    "libreoffice",
                    "--headless",
                    "--convert-to", "docx",
                    "--outdir", outdir,
                    input_path,
                ],
                check=True,
                capture_output=True,
                text=True,
            )
        except FileNotFoundError:
            raise RuntimeError(
                "LibreOffice not found. Install LibreOffice to convert .doc files."
            )
        except subprocess.CalledProcessError as exc:
            raise RuntimeError(f"LibreOffice error: {exc.stderr.strip()}")

        stem = Path(input_path).stem
        docx_path = Path(outdir) / f"{stem}.docx"
        if not docx_path.exists():
            # LibreOffice may produce a different name on some systems
            candidates = list(Path(outdir).glob("*.docx"))
            if not candidates:
                raise RuntimeError("LibreOffice produced no .docx output.")
            docx_path = candidates[0]
        return str(docx_path)
