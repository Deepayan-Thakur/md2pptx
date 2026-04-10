"""
parser.py — Markdown structure extractor.
Parses headings, content, tables, bullet lists, and numeric data.
"""
import re
from dataclasses import dataclass, field
from typing import List, Dict, Any, Optional, Tuple


@dataclass
class Table:
    title: str
    headers: List[str]
    rows: List[List[str]]
    is_numerical: bool = False

    def chart_data(self) -> Optional[Dict]:
        """Return chart-friendly data if table has a numeric column."""
        if not self.rows or not self.headers:
            return None
        numeric_col = None
        for ci in range(1, len(self.headers)):
            vals = []
            for row in self.rows:
                if ci < len(row):
                    raw = re.sub(r'[,$%\s~≈]', '', row[ci])
                    try:
                        vals.append(float(raw))
                    except ValueError:
                        break
            if len(vals) == len(self.rows):
                numeric_col = (ci, vals)
                break
        if numeric_col is None:
            return None
        ci, values = numeric_col
        labels = [row[0].strip()[:30] for row in self.rows]
        return {
            "type": "bar",
            "chart_title": self.title or self.headers[ci],
            "labels": labels[:8],
            "values": values[:8],
            "ylabel": self.headers[ci],
        }


@dataclass
class Section:
    level: int
    title: str
    content: str = ""
    bullets: List[str] = field(default_factory=list)
    subsections: List['Section'] = field(default_factory=list)
    tables: List[Table] = field(default_factory=list)
    has_numbers: bool = False

    def short_bullets(self, max_items: int = 5, max_chars: int = 150) -> List[str]:
        raw = self.bullets or [self.content]
        out = []
        import re
        for b in raw:
            b = clean_text(b).strip()
            if not b:
                continue
            sentences = [s.strip() for s in re.split(r'(?<=[.!?])\s+', b) if len(s.strip()) > 10]
            for s in sentences:
                if len(out) < max_items:
                    out.append(s)
            if len(out) >= max_items:
                break
        return out


def clean_text(text: str) -> str:
    """Strip markdown syntax, citation refs, and image tags."""
    text = re.sub(r'!\[.*?\]\(.*?\)', '', text)          # images
    text = re.sub(r'\[(\d+)\]\([^)]*\)', '', text)       # [1](url) citations
    text = re.sub(r'\[(.*?)\]\([^)]*\)', r'\1', text)    # [text](url) → text
    text = re.sub(r'\*\*(.*?)\*\*', r'\1', text)         # bold
    text = re.sub(r'\*(.*?)\*', r'\1', text)             # italic
    text = re.sub(r'`([^`]*)`', r'\1', text)             # inline code
    text = re.sub(r'#{1,6}\s', '', text)                 # headings
    text = re.sub(r'\s+', ' ', text)
    return text.strip()


def _is_table_separator(line: str) -> bool:
    return bool(re.match(r'^\s*\|?[-:\s|]+\|[-:\s|]+\|\s*$', line))


def _parse_table(lines: List[str], start: int) -> Tuple[Optional[Table], int]:
    """Parse a markdown table starting at `start`. Returns (Table, next_idx)."""
    i = start
    table_lines = []
    while i < len(lines):
        line = lines[i]
        if not line.strip().startswith('|') and not _is_table_separator(line):
            break
        if not _is_table_separator(line):
            table_lines.append(line)
        i += 1
    if len(table_lines) < 2:
        return None, start + 1

    headers = [c.strip() for c in table_lines[0].strip('| \t').split('|')]
    rows = []
    for tl in table_lines[1:]:
        cells = [c.strip() for c in tl.strip('| \t').split('|')]
        if cells and any(c for c in cells):
            rows.append(cells)

    if not rows:
        return None, i

    # Detect title from line above
    title = ""
    if start > 0:
        prev = lines[start - 1].strip()
        m = re.match(r'^(?:Title:|###?\s*|(?:\*\*)?title(?:\*\*)?:?\s*)(.*)', prev, re.I)
        if m:
            title = clean_text(m.group(1)).strip('*').strip()
        elif prev and not prev.startswith('#') and not prev.startswith('|') and not prev.startswith('-'):
            title = clean_text(prev)

    is_num = any(re.search(r'\d', ' '.join(r)) for r in rows)
    return Table(title=title, headers=headers, rows=rows, is_numerical=is_num), i


def parse_markdown(md_text: str) -> Dict[str, Any]:
    """Full markdown parse. Returns structured dict for the builder."""
    lines = md_text.split('\n')

    result: Dict[str, Any] = {
        "title": "",
        "subtitle": "",
        "executive_summary": "",
        "toc": [],            # table of contents items
        "sections": [],
        "all_tables": [],
        "has_numerical_data": False,
    }

    # ── Title (first H1) ──────────────────────────────────────────────────
    for line in lines:
        if line.startswith('# ') and not result["title"]:
            result["title"] = clean_text(line[2:])
            break

    # ── Subtitle (first ### before any H2) ────────────────────────────────
    for line in lines:
        if line.startswith('## '):
            break
        if line.startswith('### '):
            result["subtitle"] = clean_text(line[4:])
            break

    # ── Table of Contents items ───────────────────────────────────────────
    toc_items = []
    in_toc = False
    for line in lines:
        stripped = line.strip()
        if re.match(r'^\[?\s*table of contents\s*\]?', stripped, re.I):
            in_toc = True
            continue
        if in_toc:
            if stripped.startswith('## '):
                break
            m = re.match(r'^\[(\d+\.\s*[^\]]+)\]', stripped)
            if m:
                toc_items.append(clean_text(m.group(1)))
    result["toc"] = toc_items[:8]

    # ── Executive Summary ─────────────────────────────────────────────────
    exec_start = -1
    for i, line in enumerate(lines):
        if re.search(r'executive\s+summary', line, re.I) and line.startswith('#'):
            exec_start = i + 1
            break
    if exec_start > 0:
        buf = []
        for line in lines[exec_start:]:
            if line.startswith('## ') and buf:
                break
            t = clean_text(line)
            if t:
                buf.append(t)
        result["executive_summary"] = ' '.join(buf)[:800]

    # ── Section / subsection parsing ──────────────────────────────────────
    current_h2: Optional[Section] = None
    current_h3: Optional[Section] = None
    buf: List[str] = []

    def flush_h3():
        nonlocal current_h3, buf
        if current_h3 is None:
            return
        current_h3.content = ' '.join(buf).strip()
        # Extract bullet points from content
        current_h3.bullets = _extract_bullets(buf)
        buf = []
        if current_h2:
            current_h2.subsections.append(current_h3)
        current_h3 = None

    def flush_h2():
        nonlocal current_h2, buf
        flush_h3()
        if current_h2 is None:
            return
        if buf:
            current_h2.content = ' '.join(buf).strip()
            current_h2.bullets = _extract_bullets(buf)
        buf = []
        result["sections"].append(current_h2)
        current_h2 = None

    i = 0
    while i < len(lines):
        line = lines[i]
        stripped = line.strip()

        if line.startswith('## '):
            flush_h2()
            current_h2 = Section(level=2, title=clean_text(line[3:]))
        elif line.startswith('### ') and current_h2:
            flush_h3()
            current_h3 = Section(level=3, title=clean_text(line[4:]))
        elif stripped.startswith('|') or (i > 0 and lines[i-1].strip().startswith('|') and _is_table_separator(stripped)):
            tbl, i = _parse_table(lines, i)
            if tbl:
                result["all_tables"].append(tbl)
                if tbl.is_numerical:
                    result["has_numerical_data"] = True
                target = current_h3 or current_h2
                if target:
                    target.tables.append(tbl)
                    if tbl.is_numerical:
                        target.has_numbers = True
            continue
        else:
            # Accumulate text
            if stripped and not stripped.startswith('!['):
                buf.append(stripped)
                if re.search(r'\$[\d,.]+|\d+[\s]?%|\d+\s*billion|\d+\s*million', stripped, re.I):
                    result["has_numerical_data"] = True
                    for s in [current_h3, current_h2]:
                        if s:
                            s.has_numbers = True
                            break
        i += 1

    flush_h2()
    return result


def _extract_bullets(buf: List[str]) -> List[str]:
    """Extract bullet points from a text buffer."""
    bullets = []
    paragraph = []
    for line in buf:
        if re.match(r'^[-*•▸►]\s+', line) or re.match(r'^\d+\.\s+', line):
            bullets.append(re.sub(r'^[-*•▸►\d.]+\s+', '', line).strip())
        elif line.strip():
            paragraph.append(line.strip())

    if not bullets and paragraph:
        # Split long paragraphs into sentences as pseudo-bullets
        text = ' '.join(paragraph)
        sentences = re.split(r'(?<=[.!?])\s+', text)
        bullets = [s for s in sentences if len(s) > 20][:5]

    return bullets[:6]