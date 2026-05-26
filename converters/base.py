from abc import ABC, abstractmethod


class BaseConverter(ABC):
    def __init__(self, html_tables: bool = False) -> None:
        self.html_tables = html_tables

    @abstractmethod
    def convert(self, input_path: str, output_path: str, images_dir: str) -> None:
        """Convert input_path → output_path (Markdown). Images go to images_dir."""
