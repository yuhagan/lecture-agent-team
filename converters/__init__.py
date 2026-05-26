from typing import Optional

from .base import BaseConverter
from .docx_converter import DocxConverter
from .doc_converter import DocConverter
from .hwpx_converter import HwpxConverter
from .hwp_converter import HwpConverter

_REGISTRY = {
    ".docx": DocxConverter,
    ".doc": DocConverter,
    ".hwpx": HwpxConverter,
    ".hwp": HwpConverter,
}


def get_converter(ext: str, html_tables: bool = False) -> Optional[BaseConverter]:
    cls = _REGISTRY.get(ext.lower())
    return cls(html_tables=html_tables) if cls else None
