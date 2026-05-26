import subprocess

from .base import BaseConverter


class DocxConverter(BaseConverter):
    """Convert DOCX → Markdown using pandoc."""

    def convert(self, input_path: str, output_path: str, images_dir: str) -> None:
        cmd = [
            "pandoc",
            input_path,
            "-o", output_path,
            f"--extract-media={images_dir}",
            "--wrap=none",
        ]
        try:
            subprocess.run(cmd, check=True, capture_output=True, text=True)
        except FileNotFoundError:
            raise RuntimeError(
                "pandoc not found. Install: https://pandoc.org/installing.html"
            )
        except subprocess.CalledProcessError as exc:
            raise RuntimeError(f"pandoc error: {exc.stderr.strip()}")
