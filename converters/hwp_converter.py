"""
HWP (binary) → Markdown converter.

Strategy:
  1. hwp5html (pyhwp)  →  pandoc  (best quality)
  2. hwp5txt  (pyhwp)  →  write as plain-text markdown  (no pandoc needed)
  3. LibreOffice headless → DOCX → pandoc  (fallback)

chardet is used to decode hwp5txt output safely (old HWP may emit EUC-KR).
"""

import shutil
import subprocess
import sys
import tempfile
from pathlib import Path

from .base import BaseConverter
from .docx_converter import DocxConverter


def _decode_bytes(data: bytes) -> str:
    """Decode bytes, using chardet when available, otherwise UTF-8 with fallback."""
    try:
        import chardet
        detected = chardet.detect(data)
        encoding = detected.get("encoding") or "utf-8"
        return data.decode(encoding, errors="replace")
    except ImportError:
        # chardet not installed — try utf-8 then cp949 (EUC-KR superset)
        for enc in ("utf-8", "cp949", "latin-1"):
            try:
                return data.decode(enc)
            except UnicodeDecodeError:
                continue
        return data.decode("utf-8", errors="replace")


class HwpConverter(BaseConverter):

    def convert(self, input_path: str, output_path: str, images_dir: str) -> None:
        try:
            self._via_pyhwp(input_path, output_path, images_dir)
        except Exception as exc:
            print(f"  [pyhwp] {exc} — trying LibreOffice fallback…", file=sys.stderr)
            self._via_libreoffice(input_path, output_path, images_dir)

    # -----------------------------------------------------------------------
    # pyhwp path
    # -----------------------------------------------------------------------

    def _via_pyhwp(self, input_path: str, output_path: str, images_dir: str) -> None:
        if shutil.which("hwp5html"):
            try:
                self._pyhwp_html(input_path, output_path, images_dir)
                return
            except RuntimeError as exc:
                print(f"  [hwp5html] {exc} — falling back to hwp5txt…", file=sys.stderr)
                if shutil.which("hwp5txt"):
                    self._pyhwp_txt(input_path, output_path)
                    return
                raise
        if shutil.which("hwp5txt"):
            self._pyhwp_txt(input_path, output_path)
        else:
            raise RuntimeError("pyhwp CLI not found (hwp5html / hwp5txt)")

    def _pyhwp_html(self, input_path: str, output_path: str, images_dir: str) -> None:
        with tempfile.TemporaryDirectory() as tmpdir:
            result = subprocess.run(
                ["hwp5html", "--output", tmpdir, input_path],
                capture_output=True,
            )
            if result.returncode != 0:
                raise RuntimeError(
                    f"hwp5html failed: {_decode_bytes(result.stderr).strip()}"
                )

            html_files = list(Path(tmpdir).glob("*.html")) or list(Path(tmpdir).glob("*.xhtml"))
            if not html_files:
                raise RuntimeError("hwp5html produced no HTML output")

            try:
                subprocess.run(
                    [
                        "pandoc", str(html_files[0]),
                        "-o", output_path,
                        f"--extract-media={images_dir}",
                        "--wrap=none",
                    ],
                    check=True,
                    capture_output=True,
                )
            except FileNotFoundError:
                raise RuntimeError("pandoc not found")
            except subprocess.CalledProcessError as exc:
                raise RuntimeError(
                    f"pandoc failed: {_decode_bytes(exc.stderr).strip()}"
                )

    def _pyhwp_txt(self, input_path: str, output_path: str) -> None:
        result = subprocess.run(["hwp5txt", input_path], capture_output=True)
        if result.returncode != 0:
            raise RuntimeError(
                f"hwp5txt failed: {_decode_bytes(result.stderr).strip()}"
            )
        text = _decode_bytes(result.stdout)
        Path(output_path).write_text(text, encoding="utf-8")

    # -----------------------------------------------------------------------
    # LibreOffice fallback
    # -----------------------------------------------------------------------

    def _via_libreoffice(self, input_path: str, output_path: str, images_dir: str) -> None:
        with tempfile.TemporaryDirectory() as tmpdir:
            try:
                subprocess.run(
                    [
                        "libreoffice", "--headless",
                        "--convert-to", "docx",
                        "--outdir", tmpdir,
                        input_path,
                    ],
                    check=True,
                    capture_output=True,
                )
            except FileNotFoundError:
                raise RuntimeError(
                    "Neither pyhwp nor LibreOffice is available.\n"
                    "  Install pyhwp:      pip install pyhwp\n"
                    "  Install LibreOffice: https://www.libreoffice.org/"
                )
            except subprocess.CalledProcessError as exc:
                raise RuntimeError(
                    f"LibreOffice failed: {_decode_bytes(exc.stderr).strip()}"
                )

            stem = Path(input_path).stem
            docx_path = Path(tmpdir) / f"{stem}.docx"
            if not docx_path.exists():
                candidates = list(Path(tmpdir).glob("*.docx"))
                if not candidates:
                    raise RuntimeError("LibreOffice produced no .docx output.")
                docx_path = candidates[0]

            DocxConverter(html_tables=self.html_tables).convert(
                str(docx_path), output_path, images_dir
            )
