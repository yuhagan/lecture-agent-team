"""
HWPX → Markdown converter.

HWPX is a ZIP archive:
  Contents/header.xml   — style definitions (paraStyle names → heading levels)
  Contents/section*.xml — body: paragraphs, tables, inline images
  BinData/*             — embedded images
"""

import os
import re
import zipfile
import xml.etree.ElementTree as ET
from dataclasses import dataclass, field
from html import escape as html_escape
from pathlib import Path
from typing import Dict, List, Optional

from .base import BaseConverter


# ---------------------------------------------------------------------------
# XML helpers (namespace-agnostic)
# ---------------------------------------------------------------------------

def _local(tag: str) -> str:
    """'{http://...}para' → 'para'"""
    return tag.split("}", 1)[-1] if "}" in tag else tag


def _attr(elem: ET.Element, name: str, default: str = "") -> str:
    """Attribute value by local name, ignoring any namespace prefix."""
    for k, v in elem.attrib.items():
        if _local(k) == name:
            return v
    return default


def _int_attr(elem: ET.Element, *names: str) -> int:
    """Read the first matching attribute as int, defaulting to 1."""
    for name in names:
        v = _attr(elem, name)
        if v:
            try:
                return max(1, int(v))
            except ValueError:
                pass
    return 1


# ---------------------------------------------------------------------------
# Block types
# ---------------------------------------------------------------------------

@dataclass
class _Para:
    level: Optional[int]  # None = normal paragraph, 1-6 = heading
    text: str
    images: List[str] = field(default_factory=list)


@dataclass
class _Cell:
    text: str
    colspan: int = 1
    rowspan: int = 1


@dataclass
class _Table:
    rows: List[List[_Cell]]

    @property
    def has_spans(self) -> bool:
        return any(
            c.colspan > 1 or c.rowspan > 1
            for row in self.rows
            for c in row
        )


# ---------------------------------------------------------------------------
# Converter
# ---------------------------------------------------------------------------

class HwpxConverter(BaseConverter):
    """Convert HWPX (ZIP + XML) to Markdown."""

    _PARA_TAGS  = {"p", "para", "paragraph"}
    _RUN_TAGS   = {"run", "r"}
    _TEXT_TAGS  = {"t", "text"}
    _TABLE_TAGS = {"tbl", "table"}
    _ROW_TAGS   = {"tr", "tableRow", "row"}
    _CELL_TAGS  = {"tc", "tableCell", "cell"}
    _IMG_TAGS   = {"pic", "picture", "img", "image", "shapeComponent"}
    _SKIP_TAGS  = {
        "pPr", "rPr", "linesegarray", "lineSeg", "charPr",
        "paraShape", "charShape", "bullet", "numbering", "trackChange",
        "head", "foot", "footnote", "endnote",
    }

    # -----------------------------------------------------------------------
    # Public entry point
    # -----------------------------------------------------------------------

    def convert(self, input_path: str, output_path: str, images_dir: str) -> None:
        os.makedirs(images_dir, exist_ok=True)

        with zipfile.ZipFile(input_path, "r") as zf:
            names = set(zf.namelist())
            style_map = self._load_style_map(zf, names)
            img_map   = self._extract_images(zf, names, images_dir)
            blocks    = self._parse_sections(zf, names, style_map, img_map)

        md = self._render(blocks, output_path)
        Path(output_path).write_text(md, encoding="utf-8")

    # -----------------------------------------------------------------------
    # Style map  →  styleID: str → heading level 1-6 or None
    # -----------------------------------------------------------------------

    def _load_style_map(self, zf: zipfile.ZipFile, names: set) -> Dict[str, Optional[int]]:
        header = "Contents/header.xml"
        if header not in names:
            return {}
        with zf.open(header) as f:
            root = ET.fromstring(f.read())
        style_map: Dict[str, Optional[int]] = {}
        for elem in root.iter():
            if _local(elem.tag) == "paraStyle":
                sid = _attr(elem, "id")
                if sid:
                    style_map[sid] = self._heading_level(
                        _attr(elem, "engName"), _attr(elem, "name")
                    )
        return style_map

    @staticmethod
    def _heading_level(eng: str, kor: str) -> Optional[int]:
        for s in (eng, kor):
            m = re.match(r"[Hh]eading\s*(\d)", s) or re.match(r"제목\s*(\d)", s)
            if m:
                return min(int(m.group(1)), 6)
        return None

    # -----------------------------------------------------------------------
    # Image extraction  BinData/* → images_dir
    # -----------------------------------------------------------------------

    def _extract_images(
        self, zf: zipfile.ZipFile, names: set, images_dir: str
    ) -> Dict[str, str]:
        """Returns {bare_filename: abs_path} and {stem: abs_path} for ext-less refs."""
        img_map: Dict[str, str] = {}
        for name in names:
            if name.startswith("BinData/") and not name.endswith("/"):
                fname = Path(name).name
                out_path = os.path.join(images_dir, fname)
                with open(out_path, "wb") as fh:
                    fh.write(zf.read(name))
                img_map[fname] = out_path
                stem = Path(fname).stem
                if stem not in img_map:
                    img_map[stem] = out_path
        return img_map

    # -----------------------------------------------------------------------
    # Section parsing
    # -----------------------------------------------------------------------

    def _parse_sections(
        self,
        zf: zipfile.ZipFile,
        names: set,
        style_map: Dict[str, Optional[int]],
        img_map: Dict[str, str],
    ) -> List:
        section_files = sorted(
            [n for n in names if re.match(r"Contents/section\d+\.xml$", n)],
            key=lambda x: int(re.search(r"\d+", x).group()),
        )
        blocks: List = []
        for sf in section_files:
            with zf.open(sf) as f:
                root = ET.fromstring(f.read())
            self._walk(root, style_map, img_map, blocks)
        return blocks

    def _walk(
        self,
        elem: ET.Element,
        style_map: Dict[str, Optional[int]],
        img_map: Dict[str, str],
        out: List,
    ) -> None:
        tag = _local(elem.tag)
        if tag in self._PARA_TAGS:
            b = self._parse_para(elem, style_map, img_map)
            if b is not None:
                out.append(b)
        elif tag in self._TABLE_TAGS:
            b = self._parse_table(elem, img_map)
            if b is not None:
                out.append(b)
        else:
            for child in elem:
                self._walk(child, style_map, img_map, out)

    # -----------------------------------------------------------------------
    # Paragraph
    # -----------------------------------------------------------------------

    def _parse_para(
        self,
        elem: ET.Element,
        style_map: Dict[str, Optional[int]],
        img_map: Dict[str, str],
    ) -> Optional[_Para]:
        sid = _attr(elem, "styleIDRef") or _attr(elem, "styleId") or _attr(elem, "style")
        level = style_map.get(sid)

        # Fallback: explicit outline level in pPr
        if level is None:
            level = self._outline_level(elem)

        texts: List[str] = []
        images: List[str] = []
        self._extract_runs(elem, texts, images, img_map)

        text = "".join(texts).strip()
        if not text and not images:
            return None
        return _Para(level=level, text=text, images=images)

    @staticmethod
    def _outline_level(para: ET.Element) -> Optional[int]:
        for child in para:
            if _local(child.tag) in ("pPr", "parapr", "paragraphProperties"):
                for sub in child:
                    if _local(sub.tag) in ("outlineShape", "outline", "numPr"):
                        lvl = (
                            _attr(sub, "level")
                            or _attr(sub, "outlineLvl")
                            or _attr(sub, "ilvl")
                        )
                        if lvl and lvl not in ("0", "9"):
                            return min(int(lvl), 6)
        return None

    def _extract_runs(
        self,
        elem: ET.Element,
        texts: List[str],
        images: List[str],
        img_map: Dict[str, str],
    ) -> None:
        for child in elem:
            tag = _local(child.tag)
            if tag in self._SKIP_TAGS:
                continue
            if tag in self._TEXT_TAGS:
                if child.text:
                    texts.append(child.text)
            elif tag in self._IMG_TAGS:
                img = self._resolve_image(child, img_map)
                if img:
                    images.append(img)
            else:
                self._extract_runs(child, texts, images, img_map)

    def _resolve_image(self, elem: ET.Element, img_map: Dict[str, str]) -> Optional[str]:
        for attr_name in ("href", "src", "binRef", "id"):
            ref = _attr(elem, attr_name)
            if ref:
                result = self._lookup_img(ref, img_map)
                if result:
                    return result
        for child in elem:
            for attr_name in ("href", "src", "binRef", "id"):
                ref = _attr(child, attr_name)
                if ref:
                    result = self._lookup_img(ref, img_map)
                    if result:
                        return result
        return None

    @staticmethod
    def _lookup_img(ref: str, img_map: Dict[str, str]) -> Optional[str]:
        fname = Path(ref).name      # 'BIN0001.png' or just 'BIN0001'
        if fname in img_map:
            return img_map[fname]
        return img_map.get(Path(fname).stem)

    # -----------------------------------------------------------------------
    # Table  (colspan / rowspan aware)
    # -----------------------------------------------------------------------

    def _parse_table(
        self, elem: ET.Element, img_map: Dict[str, str]
    ) -> Optional[_Table]:
        rows: List[List[_Cell]] = []
        for child in elem:
            if _local(child.tag) in self._ROW_TAGS:
                row: List[_Cell] = []
                for tc in child:
                    if _local(tc.tag) in self._CELL_TAGS:
                        # colspan / rowspan — HWPX uses camelCase attributes
                        colspan = _int_attr(tc, "colSpan", "colspan", "gridSpan")
                        rowspan = _int_attr(tc, "rowSpan", "rowspan", "vMerge")

                        cell_texts: List[str] = []
                        cell_imgs: List[str] = []
                        for p in tc:
                            if _local(p.tag) in self._PARA_TAGS:
                                t: List[str] = []
                                self._extract_runs(p, t, cell_imgs, img_map)
                                cell_texts.append("".join(t).strip())
                        text = " ".join(filter(None, cell_texts))
                        row.append(_Cell(text=text, colspan=colspan, rowspan=rowspan))
                if row:
                    rows.append(row)
        return _Table(rows) if rows else None

    # -----------------------------------------------------------------------
    # Markdown rendering
    # -----------------------------------------------------------------------

    def _render(self, blocks: List, output_path: str) -> str:
        output_dir = os.path.dirname(os.path.abspath(output_path))
        lines: List[str] = []

        for block in blocks:
            if isinstance(block, _Para):
                if block.level:
                    lines.append(f"{'#' * block.level} {block.text}")
                elif block.text:
                    lines.append(block.text)
                for img in block.images:
                    rel = _rel_path(img, output_dir)
                    lines.append(f"![{Path(img).name}]({rel})")
                lines.append("")

            elif isinstance(block, _Table):
                # Auto-detect: use HTML if spans present or --html-tables requested
                if self.html_tables or block.has_spans:
                    lines.append(_render_table_html(block.rows))
                else:
                    lines.append(_render_table_md(block.rows))
                lines.append("")

        return "\n".join(lines).strip() + "\n"


# ---------------------------------------------------------------------------
# Table renderers (module-level to keep the class clean)
# ---------------------------------------------------------------------------

def _render_table_md(rows: List[List[_Cell]]) -> str:
    if not rows:
        return ""
    ncols = max(len(r) for r in rows)

    def pad(row: List[_Cell]) -> List[str]:
        texts = [c.text for c in row]
        return texts + [""] * (ncols - len(texts))

    def fmt(texts: List[str]) -> str:
        return "| " + " | ".join(t.replace("|", "\\|") for t in texts) + " |"

    out = [fmt(pad(rows[0])), "| " + " | ".join(["---"] * ncols) + " |"]
    out.extend(fmt(pad(r)) for r in rows[1:])
    return "\n".join(out)


def _render_table_html(rows: List[List[_Cell]]) -> str:
    if not rows:
        return ""
    out = ["<table>"]
    for i, row in enumerate(rows):
        out.append("  <tr>")
        tag = "th" if i == 0 else "td"
        for cell in row:
            attrs = ""
            if cell.colspan > 1:
                attrs += f' colspan="{cell.colspan}"'
            if cell.rowspan > 1:
                attrs += f' rowspan="{cell.rowspan}"'
            out.append(f"    <{tag}{attrs}>{html_escape(cell.text)}</{tag}>")
        out.append("  </tr>")
    out.append("</table>")
    return "\n".join(out)


# ---------------------------------------------------------------------------
# Path helper
# ---------------------------------------------------------------------------

def _rel_path(abs_img: str, output_dir: str) -> str:
    try:
        return os.path.relpath(abs_img, output_dir).replace("\\", "/")
    except ValueError:
        return abs_img.replace("\\", "/")
