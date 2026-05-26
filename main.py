#!/usr/bin/env python3
"""doc2md — Convert HWP/HWPX/DOC/DOCX files to Markdown."""

import argparse
import io
import sys
from pathlib import Path
from typing import List

# Windows 터미널이 cp949일 때도 유니코드 출력 가능하게 재설정
if sys.stdout.encoding and sys.stdout.encoding.lower() not in ("utf-8", "utf-16"):
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")
if sys.stderr.encoding and sys.stderr.encoding.lower() not in ("utf-8", "utf-16"):
    sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding="utf-8", errors="replace")

from cleanup import cleanup_file
from converters import get_converter
from utils import check_tools, print_tool_errors

SUPPORTED_EXTS = {".hwpx", ".hwp", ".docx", ".doc"}


# ---------------------------------------------------------------------------
# Single-file conversion
# ---------------------------------------------------------------------------

def convert_one(
    input_path: Path,
    output_path: Path,
    images_dir: Path,
    html_tables: bool,
) -> bool:
    """Convert a single file. Returns True on success."""
    ext = input_path.suffix.lower()
    errors = check_tools(ext)
    if errors:
        print(f"\n  ✗ Cannot convert {input_path.name}:", file=sys.stderr)
        print_tool_errors(errors, fatal=True)
        return False

    converter = get_converter(ext, html_tables=html_tables)
    if converter is None:
        print(f"  ✗ Unsupported format: {ext}", file=sys.stderr)
        return False

    try:
        converter.convert(str(input_path), str(output_path), str(images_dir))
        cleanup_file(output_path)
        print(f"  ✓  {input_path}  →  {output_path}")
        return True
    except Exception as exc:
        print(f"  ✗  {input_path}: {exc}", file=sys.stderr)
        return False


# ---------------------------------------------------------------------------
# Batch conversion
# ---------------------------------------------------------------------------

def convert_batch(
    input_dir: Path,
    output_dir: Path,
    images_dir: Path,
    html_tables: bool,
    recursive: bool,
) -> None:
    pattern = "**/*" if recursive else "*"
    files: List[Path] = sorted(
        f for f in input_dir.glob(pattern)
        if f.is_file() and f.suffix.lower() in SUPPORTED_EXTS
    )

    if not files:
        print(f"No supported files found in {input_dir}", file=sys.stderr)
        return

    print(f"Found {len(files)} file(s) in '{input_dir}':\n")
    ok = fail = 0

    for src in files:
        rel = src.relative_to(input_dir)
        dst = (output_dir / rel).with_suffix(".md")
        dst.parent.mkdir(parents=True, exist_ok=True)

        if convert_one(src, dst, images_dir, html_tables):
            ok += 1
        else:
            fail += 1

    print(f"\nDone — {ok} succeeded, {fail} failed.")


# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------

def main() -> None:
    parser = argparse.ArgumentParser(
        prog="doc2md",
        description="Convert HWP/HWPX/DOC/DOCX to Markdown",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=(
            "Examples:\n"
            "  python main.py report.hwpx -o report.md\n"
            "  python main.py ./docs_folder/ -o ./output/ --recursive\n"
            "  python main.py table-heavy.hwpx --html-tables"
        ),
    )
    parser.add_argument(
        "input",
        help="Input file (.hwpx .hwp .docx .doc) or directory for batch mode",
    )
    parser.add_argument(
        "-o", "--output",
        help=(
            "Output .md path (single file) or output directory (batch). "
            "Default: same location as input with .md extension"
        ),
    )
    parser.add_argument(
        "--images-dir",
        default="output/images",
        metavar="DIR",
        help="Directory for extracted images (default: output/images)",
    )
    parser.add_argument(
        "--html-tables",
        action="store_true",
        help="Render all tables as HTML <table> instead of Markdown pipes",
    )
    parser.add_argument(
        "--recursive", "-r",
        action="store_true",
        help="(Batch mode) Recurse into subdirectories",
    )

    args = parser.parse_args()

    input_path = Path(args.input)
    if not input_path.exists():
        sys.exit(f"Error: path not found: {input_path}")

    images_dir = Path(args.images_dir)
    images_dir.mkdir(parents=True, exist_ok=True)

    # ── Batch mode ────────────────────────────────────────────────────────
    if input_path.is_dir():
        output_dir = Path(args.output) if args.output else input_path.parent / "output"
        output_dir.mkdir(parents=True, exist_ok=True)
        convert_batch(input_path, output_dir, images_dir, args.html_tables, args.recursive)
        return

    # ── Single file ───────────────────────────────────────────────────────
    if input_path.suffix.lower() not in SUPPORTED_EXTS:
        sys.exit(
            f"Error: unsupported format '{input_path.suffix}'. "
            f"Supported: {' '.join(sorted(SUPPORTED_EXTS))}"
        )

    output_path = Path(args.output) if args.output else input_path.with_suffix(".md")
    output_path.parent.mkdir(parents=True, exist_ok=True)

    success = convert_one(input_path, output_path, images_dir, args.html_tables)
    if not success:
        sys.exit(1)


if __name__ == "__main__":
    main()
