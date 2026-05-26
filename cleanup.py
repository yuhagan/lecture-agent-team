"""cleanup.py — pandoc xhtml 변환 결과 후처리.

제거 대상:
  - [텍스트]{.attrs}  → 텍스트 (span unwrap, 중첩 포함)
  - []{.attrs}        → 삭제 (빈 span, multiline attrs 포함)
  - :::+ {.attrs} 줄 → 삭제 (div fence)
  - 연속 빈 줄 3개+  → 빈 줄 1개로 축소
"""

import re
from pathlib import Path


def _remove_empty_spans(text: str) -> str:
    """[]{...} 형태의 빈 span 제거. 중첩 해소를 위해 반복."""
    pattern = re.compile(r'\[\]\{[^{}]*\}', re.DOTALL)
    for _ in range(10):
        new = pattern.sub('', text)
        if new == text:
            break
        text = new
    return text


def _unwrap_spans(text: str) -> str:
    """[content]{.attrs} → content. 중첩 해소를 위해 반복."""
    pattern = re.compile(r'\[([^\[\]]+?)\]\{[^{}]*\}', re.DOTALL)
    for _ in range(10):
        new = pattern.sub(r'\1', text)
        if new == text:
            break
        text = new
    return text


def _remove_div_fences(text: str) -> str:
    """:::+ 로 시작하는 div fence 줄 제거."""
    return re.sub(r'^:{2,}[^\n]*\n?', '', text, flags=re.MULTILINE)


def _collapse_blank_lines(text: str) -> str:
    return re.sub(r'\n{3,}', '\n\n', text)


def cleanup(text: str) -> str:
    text = _remove_empty_spans(text)
    text = _unwrap_spans(text)
    text = _remove_div_fences(text)
    text = _collapse_blank_lines(text)
    return text.strip() + '\n'


def cleanup_file(path: str | Path) -> None:
    p = Path(path)
    original = p.read_text(encoding='utf-8')
    cleaned = cleanup(original)
    p.write_text(cleaned, encoding='utf-8')


if __name__ == '__main__':
    import sys
    if len(sys.argv) < 2:
        print(f'Usage: python cleanup.py <file.md>', file=sys.stderr)
        sys.exit(1)
    for arg in sys.argv[1:]:
        cleanup_file(arg)
        print(f'  cleaned: {arg}')
