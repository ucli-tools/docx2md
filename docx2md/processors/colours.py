"""
Table colours: a cell's background and the colour of its text, kept from the Word document.

Pandoc drops both. A book marks its tables with them (a prime value in red, a block of a matrix
shaded light blue) and its text names those colours, so they are kept, in a form every tool reads:

- a background, as raw TeX at the cell's start: `` `\\cellcolor[HTML]{BDD6EE}`{=latex} ``;
- a text colour, around the cell's content: ``$\\textcolor[HTML]{FF0000}{...}$``.

mdtexpdf draws both in the PDF (and gives the EPUB the same colours), the school's import shows
both, mdaudiobook reads neither aloud. A cell whose text mixes colours, or mixes text with maths
in a colour, keeps its background only, and is counted.

Each Word table is matched with the Markdown table pandoc made from it by its shape and by the
digits of its cells, so a table pandoc did not make (a layout table) is passed over, never
mismatched. A table given colours is written as a pipe table.
"""

import re
import zipfile
from dataclasses import dataclass, field
from pathlib import Path
from typing import Dict, List, Optional, Tuple, Union
from xml.etree import ElementTree as ET

from docx2md.utils.logging_utils import get_logger

logger = get_logger(__name__)

_W = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
_M = '{http://schemas.openxmlformats.org/officeDocument/2006/math}'
_BLACK = {None, 'auto', '000000'}
MIXED = 'mixed'


@dataclass
class Cell:
    fill: Optional[str] = None       # RRGGBB, or None
    colour: Optional[str] = None     # RRGGBB, None (black), or MIXED
    digits: str = ''


@dataclass
class WordTable:
    rows: List[List[Cell]]
    merged: bool = False
    container: bool = False          # its cells hold other tables (a layout table)

    @property
    def shape(self) -> Tuple[int, int]:
        return len(self.rows), max((len(r) for r in self.rows), default=0)

    @property
    def coloured(self) -> bool:
        return any(c.fill or c.colour for r in self.rows for c in r)


def _cell(tc: ET.Element) -> Cell:
    pr = tc.find(f'{_W}tcPr')
    fill = None
    if pr is not None:
        shd = pr.find(f'{_W}shd')
        value = shd.get(f'{_W}fill') if shd is not None else None
        if value and value.lower() != 'auto' and value.upper() != 'FFFFFF':
            fill = value.upper()
    colours = set()
    text = []
    for run in tc.iter():
        if run.tag not in (f'{_W}r', f'{_M}r'):
            continue
        words = ''.join(t.text or '' for t in run if t.tag in (f'{_W}t', f'{_M}t'))
        if not words.strip():
            continue
        text.append(words)
        c = run.find(f'{_W}rPr/{_W}color')
        value = c.get(f'{_W}val') if c is not None else None
        colours.add(None if value in _BLACK else value.upper())
    colour = None
    if len(colours) == 1:
        colour = next(iter(colours))
    elif len(colours) > 1:
        colour = MIXED
    return Cell(fill, colour, re.sub(r'\D', '', ''.join(text)))


def word_tables(docx: Union[str, Path]) -> List[WordTable]:
    """The document's tables in order, a table set inside a cell after its container; each
    cell's background, text colour and digits."""
    with zipfile.ZipFile(docx) as z:
        body = ET.fromstring(z.read('word/document.xml')).find(f'{_W}body')
    tables: List[WordTable] = []

    def walk(parent: ET.Element) -> None:
        for el in parent:
            if el.tag == f'{_W}tbl':
                own = [tr for tr in el.findall(f'{_W}tr')]
                merged = any(e.tag in (f'{_W}gridSpan', f'{_W}vMerge')
                             for tr in own for tc in tr.findall(f'{_W}tc') for e in tc.findall(f'{_W}tcPr/*'))
                rows = [[_cell(tc) for tc in tr.findall(f'{_W}tc')] for tr in own]
                container = any(tc.find(f'.//{_W}tbl') is not None for tr in own for tc in tr.findall(f'{_W}tc'))
                tables.append(WordTable(rows, merged, container))
                # A table set inside a cell (pandoc writes it out after its container) follows it.
                for tr in own:
                    for tc in tr.findall(f'{_W}tc'):
                        walk(tc)
            elif el.tag == f'{_W}sdt':
                content = el.find(f'{_W}sdtContent')
                if content is not None:
                    walk(content)

    walk(body)
    return tables


# ---------------------------------------------------------------------------
# The Markdown's tables
# ---------------------------------------------------------------------------

_PIPE_RULE = re.compile(r'^\s*\|?\s*:?-+:?\s*(?:\|\s*:?-+:?\s*)+\|?\s*$')
_RULE = re.compile(r'^\s*-{3,}\s*$')
_COLUMNS = re.compile(r'^\s*-+(?: +-+)+\s*$')


@dataclass
class MdTable:
    start: int                       # first line
    end: int                         # one past the last line
    rows: List[List[str]] = field(default_factory=list)

    @property
    def shape(self) -> Tuple[int, int]:
        return len(self.rows), max((len(r) for r in self.rows), default=0)


def pipe_cells(line: str) -> List[str]:
    """A pipe table's row as its cells, split on the « | » outside maths."""
    s = line.strip()
    if s.startswith('|'):
        s = s[1:]
    if s.endswith('|') and not s.endswith('\\|'):
        s = s[:-1]
    out, cell, math, k = [], '', False, 0
    while k < len(s):
        c = s[k]
        if c == '\\' and k + 1 < len(s):
            cell += s[k:k + 2]
            k += 2
            continue
        if c == '$':
            math = not math
        if c == '|' and not math:
            out.append(cell.strip())
            cell = ''
        else:
            cell += c
        k += 1
    out.append(cell.strip())
    return out


def md_tables(lines: List[str]) -> List[MdTable]:
    """The pipe tables and pandoc's multiline and simple tables, in order."""
    found: List[MdTable] = []
    i = 0
    while i < len(lines):
        line = lines[i]
        nxt = lines[i + 1] if i + 1 < len(lines) else ''
        if line.strip().startswith('|') and _PIPE_RULE.match(nxt) and '|' in nxt:
            n = i + 2
            while n < len(lines) and lines[n].strip().startswith('|'):
                n += 1
            rows = [pipe_cells(line)] + [pipe_cells(ln) for ln in lines[i + 2:n]]
            found.append(MdTable(i, n, rows))
            i = n
            continue
        multi = _RULE.match(line) and any(_COLUMNS.match(ln) for ln in lines[i + 1:i + 4])
        simple = not multi and _COLUMNS.match(nxt) and line.strip() and not _RULE.match(line)
        if multi or simple:
            sep = next(n for n in range(i + 1, len(lines)) if _COLUMNS.match(lines[n])) if multi else i + 1
            cols = [(m.start(), m.end()) for m in re.finditer(r'-+', lines[sep])]

            def cells(block: List[str]) -> List[str]:
                out = []
                for c, (a, _) in enumerate(cols):
                    b = cols[c + 1][0] if c + 1 < len(cols) else None
                    out.append(' '.join(ln[a:b].strip() for ln in block if ln[a:b].strip()))
                return out

            head = lines[i + 1:sep] if multi else [line]
            rows, row, n = [], [], sep + 1
            while n < len(lines):
                if (multi and _RULE.match(lines[n])) or (not multi and not lines[n].strip()):
                    break
                if not lines[n].strip():
                    if row:
                        rows.append(row)
                    row = []
                else:
                    row.append(lines[n])
                n += 1
            if row:
                rows.append(row)
            if not multi:
                rows = [[r] for r in sum(rows, [])]
            table = MdTable(i, n + 1 if multi else n, [cells(head)] + [cells(r) for r in rows])
            found.append(table)
            i = table.end
            continue
        i += 1
    return found


_MARKUP = re.compile(r'`\\cellcolor\[HTML\]\{[0-9A-Fa-f]{6}\}`\{=latex\}|\\textcolor\[HTML\]\{[0-9A-Fa-f]{6}\}')


def _digits(md_cell: str) -> str:
    t = re.sub(r'\\[A-Za-z]+', ' ', _MARKUP.sub('', md_cell))   # colours kept, commands (\frac, \pi)
    return re.sub(r'\D', '', t)


def _same(word: WordTable, md: MdTable) -> bool:
    if word.shape != md.shape:
        return False
    cells = [(w.digits, _digits(m)) for wr, mr in zip(word.rows, md.rows) for w, m in zip(wr, mr)]
    agree = sum(1 for a, b in cells if a == b)
    return agree >= 0.9 * len(cells)


# ---------------------------------------------------------------------------
# Colouring a cell
# ---------------------------------------------------------------------------

_ONE_MATH = re.compile(r'^\$((?:[^$\\]|\\.)+)\$$')


def coloured_cell(text: str, cell: Cell, skipped: Dict[str, int]) -> str:
    """A Markdown cell with its background and its text colour (a cell that has them already
    keeps them as they are)."""
    body = text.strip()
    fill = re.match(r'^`\\cellcolor\[HTML\]\{[0-9A-Fa-f]{6}\}`\{=latex\}', body)
    if fill:
        body = body[fill.end():]
    if '\\textcolor[HTML]' in body:
        pass
    elif cell.colour and cell.colour != MIXED and body:
        bold = body.startswith('**') and body.endswith('**') and len(body) > 4
        inner = body[2:-2] if bold else body
        m = _ONE_MATH.match(inner)
        if m:
            inner = f'$\\textcolor[HTML]{{{cell.colour}}}{{{m.group(1)}}}$'
        elif '$' not in inner and re.fullmatch(r'[\d.,\s−-]+', inner):
            inner = f'$\\textcolor[HTML]{{{cell.colour}}}{{{inner}}}$'
        elif '$' not in inner:
            words = inner.replace('\\', '\\textbackslash{}').replace('{', '\\{').replace('}', '\\}')
            inner = f'$\\textcolor[HTML]{{{cell.colour}}}{{\\text{{{words}}}}}$'
        else:
            skipped['text and maths in one colour'] = skipped.get('text and maths in one colour', 0) + 1
        body = f'**{inner}**' if bold else inner
    elif cell.colour == MIXED:
        skipped['colours mixed in one cell'] = skipped.get('colours mixed in one cell', 0) + 1
    if cell.fill:
        body = f'`\\cellcolor[HTML]{{{cell.fill}}}`{{=latex}}{body}'
    return body


def _pipe(rows: List[List[str]]) -> List[str]:
    width = max(len(r) for r in rows)
    rows = [r + [''] * (width - len(r)) for r in rows]

    def line(r: List[str]) -> str:
        # A bar inside maths would end the cell: it becomes \vert.
        safe = [re.sub(r'\$(?:[^$\\]|\\.)*\$', lambda m: m.group(0).replace('|', '\\vert '), c) for c in r]
        return '| ' + ' | '.join(safe) + ' |'

    return [line(rows[0]), '|' + '---|' * width] + [line(r) for r in rows[1:]]


def keep_table_colours(content: str, docx: Union[str, Path]) -> Tuple[str, Dict[str, int]]:
    """The Markdown with its tables' colours from the Word document, and what was done."""
    lines = content.split('\n')
    words = word_tables(docx)
    tables = md_tables(lines)
    stats = {'tables_coloured': 0, 'cells_filled': 0, 'cells_coloured': 0}
    skipped: Dict[str, int] = {}
    replacements = []
    w = 0
    for md in tables:
        j = next((k for k in range(w, len(words)) if _same(words[k], md)), None)
        if j is None:
            continue
        w = j + 1
        word = words[j]
        if not word.coloured:
            continue
        if word.merged:
            skipped['tables with merged cells'] = skipped.get('tables with merged cells', 0) + 1
            continue
        rows = [[coloured_cell(text, cell, skipped) for text, cell in zip(mr, wr)]
                for mr, wr in zip(md.rows, word.rows)]
        stats['tables_coloured'] += 1
        stats['cells_filled'] += sum(1 for r in word.rows for c in r if c.fill)
        stats['cells_coloured'] += sum(1 for r in word.rows for c in r if c.colour and c.colour != MIXED)
        replacements.append((md.start, md.end, _pipe(rows)))
    for start, end, new in reversed(replacements):
        lines[start:end] = new
    for what, n in skipped.items():
        logger.warning(f'table colours not kept: {what}: {n}')
    stats['colours_skipped'] = sum(skipped.values())
    return '\n'.join(lines), stats


def colour_counts(docx: Union[str, Path], markdown: str) -> Tuple[Tuple[int, int], Tuple[int, int], List[str]]:
    """What `docx2md verify` compares: the Word tables' cell backgrounds and one-colour cells
    (layout tables aside), and the backgrounds and colours in the Markdown's tables."""
    words = [t for t in word_tables(docx) if not t.container]
    kept = [t for t in words if not t.merged]
    want = (sum(1 for t in kept for r in t.rows for c in r if c.fill),
            sum(1 for t in kept for r in t.rows for c in r if c.colour and c.colour != MIXED))
    lines = markdown.split('\n')
    table_lines = [line for t in md_tables(lines) for line in lines[t.start:t.end]]
    got = (sum(line.count('\\cellcolor[HTML]') for line in table_lines),
           sum(line.count('\\textcolor[HTML]') for line in table_lines))
    notes = []
    merged = len(words) - len(kept)
    if merged:
        notes.append(f'{merged} table(s) with merged cells keep no colour')
    mixed = sum(1 for t in kept for r in t.rows for c in r if c.colour == MIXED)
    if mixed:
        notes.append(f'{mixed} cell(s) mix colours in their text and keep their background only')
    return want, got, notes
