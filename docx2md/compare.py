"""
Compare two Word documents, and change one exactly.

Made for correcting a book's original: `compare` lists every difference between
the document before and after a correction (where it is, and the whole passage
before and after), so the comparison proves that the change is the one decided
and nothing else; `edit` makes a small correction (a value, a word) in one
place, leaving every other byte of the document as it was.

Both read the .docx XML directly. A paragraph is one passage; a table row is
one passage, its cells kept apart, so a change says which cell it is in.
"""

import hashlib
import io
import re
import subprocess
import zipfile
from dataclasses import dataclass, field
from difflib import SequenceMatcher
from pathlib import Path
from typing import BinaryIO, Dict, List, Optional, Tuple, Union
from xml.etree import ElementTree as ET

_W = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
_M = '{http://schemas.openxmlformats.org/officeDocument/2006/math}'
_DOCUMENT = 'word/document.xml'

Source = Union[str, Path, bytes, BinaryIO]

# Attributes Word rewrites on every save without changing anything a reader sees.
_NOISE = re.compile(r'\s[\w.-]+:(?:rsid\w*|paraId|textId)="[^"]*"')


def _zip(src: Source) -> zipfile.ZipFile:
    if isinstance(src, (bytes, bytearray)):
        src = io.BytesIO(src)
    return zipfile.ZipFile(src)


def _entries(src: Source) -> Dict[str, bytes]:
    with _zip(src) as z:
        return {name: z.read(name) for name in z.namelist()}


# ---------------------------------------------------------------------------
# Reading a passage as text
# ---------------------------------------------------------------------------

def _outline_levels(styles_xml: Optional[bytes]) -> Dict[str, int]:
    """Each paragraph style's heading level (0 = heading 1), following basedOn."""
    if not styles_xml:
        return {}
    root = ET.fromstring(styles_xml)
    own: Dict[str, Optional[int]] = {}
    based: Dict[str, str] = {}
    for st in root.iter(f'{_W}style'):
        sid = st.get(f'{_W}styleId')
        lvl = st.find(f'{_W}pPr/{_W}outlineLvl')
        own[sid] = int(lvl.get(f'{_W}val')) if lvl is not None else None
        b = st.find(f'{_W}basedOn')
        if b is not None:
            based[sid] = b.get(f'{_W}val')
    levels = {}
    for sid in own:
        s, seen = sid, set()
        while s is not None and s not in seen:
            seen.add(s)
            if own.get(s) is not None:
                levels[sid] = own[s]
                break
            s = based.get(s)
    return levels


def _simple(t: str) -> bool:
    return bool(re.fullmatch(r'[\w.πε∞]+', t))


def _group(t: str) -> str:
    return t if _simple(t) else f'({t})'


def _math(el: ET.Element) -> str:
    """OMML as a line of text: 2iπ/13, x^2, √(a+b), e^(2iπ/3)."""
    tag = el.tag

    def kid(name: str) -> Optional[ET.Element]:
        return el.find(f'{_M}{name}')

    def inner(name: str) -> str:
        k = kid(name)
        return _math(k) if k is not None else ''

    if tag == f'{_M}t':
        return el.text or ''
    if tag == f'{_M}f':
        return f"{_group(inner('num'))}/{_group(inner('den'))}"
    if tag == f'{_M}sSup':
        return f"{inner('e')}^{_group(inner('sup'))}"
    if tag == f'{_M}sSub':
        return f"{inner('e')}_{_group(inner('sub'))}"
    if tag == f'{_M}sSubSup':
        return f"{inner('e')}_{_group(inner('sub'))}^{_group(inner('sup'))}"
    if tag == f'{_M}rad':
        deg = inner('deg')
        return f"{'(' + deg + ')' if deg else ''}√{_group(inner('e'))}"
    if tag == f'{_M}d':
        pr = kid('dPr')
        beg, end, sep = '(', ')', '|'
        if pr is not None:
            for name, default in (('begChr', '('), ('endChr', ')'), ('sepChr', '|')):
                c = pr.find(f'{_M}{name}')
                val = c.get(f'{_M}val') if c is not None else default
                if name == 'begChr':
                    beg = val
                elif name == 'endChr':
                    end = val
                else:
                    sep = val
        return beg + sep.join(_math(e) for e in el.findall(f'{_M}e')) + end
    if tag == f'{_M}nary':
        pr = kid('naryPr')
        c = pr.find(f'{_M}chr') if pr is not None else None
        op = c.get(f'{_M}val') if c is not None else '∫'
        sub, sup = inner('sub'), inner('sup')
        return f"{op}{'_' + _group(sub) if sub else ''}{'^' + _group(sup) if sup else ''} {inner('e')}"
    if tag.endswith('Pr'):
        return ''
    return ''.join(_math(c) for c in el)


def _text(el: ET.Element) -> str:
    """A paragraph's words, its maths as lines of text, in the order a reader meets them."""
    out = []

    def walk(e: ET.Element) -> None:
        if e.tag == f'{_M}oMath':
            out.append(_math(e))
            return
        if e.tag == f'{_W}t':
            out.append(e.text or '')
        elif e.tag == f'{_W}tab':
            out.append('\t')
        elif e.tag in (f'{_W}br', f'{_W}cr'):
            out.append(' ')
        elif e.tag in (f'{_W}instrText', f'{_W}delText'):
            return
        for c in e:
            walk(c)

    walk(el)
    return ''.join(out)


# ---------------------------------------------------------------------------
# The document as passages
# ---------------------------------------------------------------------------

@dataclass
class Passage:
    kind: str                      # 'paragraph' or 'row'
    section: Tuple[str, ...]       # the headings above it, outermost first
    text: str                      # a row: its cells joined by ' | '
    cells: Tuple[str, ...] = ()
    table: int = 0                 # a row: its table, counted from 1 in the document
    row: int = 0                   # a row: counted from 1
    xml: bytes = b''

    @property
    def signature(self) -> str:
        return hashlib.sha256(_NOISE.sub('', self.xml.decode('utf-8')).encode('utf-8')).hexdigest()


def _passages(document_xml: bytes, levels: Dict[str, int]) -> Tuple[List[Passage], Dict[int, str]]:
    """The body's paragraphs and table rows in order, and each table's caption (the first
    paragraph beginning « Table » among the three after it)."""
    body = ET.fromstring(document_xml).find(f'{_W}body')
    passages: List[Passage] = []
    captions: Dict[int, str] = {}
    section: List[str] = []
    tables = 0
    after_table: List[int] = []    # [table, paragraphs seen since]

    def heading(p: ET.Element) -> Optional[int]:
        pr = p.find(f'{_W}pPr')
        if pr is None:
            return None
        lvl = pr.find(f'{_W}outlineLvl')
        if lvl is not None:
            return int(lvl.get(f'{_W}val'))
        st = pr.find(f'{_W}pStyle')
        return levels.get(st.get(f'{_W}val')) if st is not None else None

    def walk(parent: ET.Element) -> None:
        nonlocal tables, after_table
        for el in parent:
            if el.tag == f'{_W}p':
                text = _text(el)
                lvl = heading(el)
                if lvl is not None and lvl < 9 and text.strip():
                    del section[lvl:]
                    section.extend([''] * (lvl - len(section)))
                    section.append(text.strip())
                passages.append(Passage('paragraph', tuple(s for s in section if s), text, xml=ET.tostring(el)))
                if after_table:
                    t, seen = after_table
                    if text.strip().startswith('Table') and t not in captions:
                        captions[t] = text.strip()
                        after_table = []
                    elif text.strip():
                        after_table = [t, seen + 1] if seen < 2 else []
            elif el.tag == f'{_W}tbl':
                tables += 1
                for r, tr in enumerate(el.findall(f'{_W}tr'), start=1):
                    cells = tuple(' / '.join(_text(p).strip() for p in tc.iter(f'{_W}p') if _text(p).strip())
                                  for tc in tr.findall(f'{_W}tc'))
                    passages.append(Passage('row', tuple(s for s in section if s), ' | '.join(cells), cells,
                                            tables, r, ET.tostring(tr)))
                after_table = [tables, 0]
            elif el.tag == f'{_W}sdt':
                content = el.find(f'{_W}sdtContent')
                if content is not None:
                    walk(content)

    walk(body)
    return passages, captions


# ---------------------------------------------------------------------------
# The comparison
# ---------------------------------------------------------------------------

@dataclass
class Change:
    kind: str                       # 'changed', 'added' or 'removed'
    where: str
    before: str
    after: str
    details: List[str] = field(default_factory=list)
    passage: Optional[Passage] = None


@dataclass
class Comparison:
    parts: List[Tuple[str, str]]    # (entry, 'added' | 'removed' | 'changed') other than the text
    changes: List[Change]
    passages: int


def _where(p: Passage, captions: Dict[int, str]) -> str:
    place = ' › '.join(p.section) or '(before the first heading)'
    if p.kind == 'row':
        cap = captions.get(p.table)
        name = f"table {p.table} of the document" + (f" ({cap.split('–')[0].split(' - ')[0].strip()})" if cap else '')
        return f'{place} › {name}, row {p.row}'
    return place


def _cell_details(a: Passage, b: Passage) -> List[str]:
    if a.kind != 'row' or b.kind != 'row':
        return []
    if len(a.cells) != len(b.cells):
        return [f'cells: {len(a.cells)} → {len(b.cells)}']
    return [f'cell {i}: « {x} » → « {y} »' for i, (x, y) in enumerate(zip(a.cells, b.cells), start=1) if x != y]


def compare(old: Source, new: Source) -> Comparison:
    """Every difference between two Word documents: the package's other parts that differ,
    and each passage changed, added or removed, with where it is and its whole text before
    and after. A passage whose text is the same but whose formatting is not (a colour, a
    style) is a change too."""
    a, b = _entries(old), _entries(new)
    parts = [(n, 'removed') for n in sorted(set(a) - set(b))] + [(n, 'added') for n in sorted(set(b) - set(a))]
    parts += [(n, 'changed') for n in sorted(set(a) & set(b)) if n != _DOCUMENT and a[n] != b[n]]
    pa, ca = _passages(a[_DOCUMENT], _outline_levels(a.get('word/styles.xml')))
    pb, cb = _passages(b[_DOCUMENT], _outline_levels(b.get('word/styles.xml')))
    changes: List[Change] = []
    matcher = SequenceMatcher(None, [p.signature for p in pa], [p.signature for p in pb], autojunk=False)
    for op, i1, i2, j1, j2 in matcher.get_opcodes():
        if op == 'equal':
            continue
        olds, news = pa[i1:i2], pb[j1:j2]
        if op == 'replace' and len(olds) == len(news):
            for x, y in zip(olds, news):
                details = _cell_details(x, y)
                if x.text == y.text:
                    details.append('the same words, its formatting changed (a colour, a style, a font)')
                changes.append(Change('changed', _where(y, cb), x.text, y.text, details, y))
            continue
        changes += [Change('removed', _where(x, ca), x.text, '', passage=x) for x in olds]
        changes += [Change('added', _where(y, cb), '', y.text, passage=y) for y in news]
    return Comparison(parts, changes, len(pb))


# ---------------------------------------------------------------------------
# Where in the PDF
# ---------------------------------------------------------------------------

def pdf_pages(pdf: Union[str, Path]) -> List[str]:
    """Each page's text, as pdftotext reads it."""
    out = subprocess.run(['pdftotext', '-layout', str(pdf), '-'], capture_output=True, check=True).stdout
    return out.decode('utf-8', 'replace').split('\f')


def page_of(change: Change, pages: List[str], captions_or_heading: Optional[str] = None) -> Optional[Tuple[int, str]]:
    """The first PDF page (counted from 1) holding the passage's own words, else its table's
    caption, else its section's heading; with what was found there."""
    flat = [re.sub(r'\s+', ' ', p) for p in pages]
    text = change.before or change.after
    words = re.findall(r'[A-Za-z][A-Za-z\'’-]+', text)
    keys = []
    if len(words) >= 6:
        keys.append((' '.join(words[:6]), 'its words'))
    if captions_or_heading:
        keys.append((captions_or_heading, 'its caption'))
    if change.passage and change.passage.section:
        keys.append((change.passage.section[-1], 'its section heading'))
    for key, what in keys:
        pattern = r'\s*'.join(re.escape(w) for w in key.split())
        for n, page in enumerate(flat, start=1):
            if re.search(pattern, page):
                return n, what
    return None


# ---------------------------------------------------------------------------
# The report
# ---------------------------------------------------------------------------

def sha256(src: Union[str, Path]) -> str:
    h = hashlib.sha256()
    with open(src, 'rb') as f:
        for block in iter(lambda: f.read(1 << 20), b''):
            h.update(block)
    return h.hexdigest()


def report(cmp: Comparison, old: Union[str, Path], new: Union[str, Path], pdf: Optional[Union[str, Path]] = None) -> str:
    """The comparison as Markdown, for a change log."""
    lines = [
        f'# {Path(old).name} → {Path(new).name}',
        '',
        f'- Before: `{Path(old).name}`, SHA-256 `{sha256(old)}`',
        f'- After: `{Path(new).name}`, SHA-256 `{sha256(new)}`',
        f'- Passages compared: {cmp.passages}; changed, added or removed: {len(cmp.changes)}',
        f"- The package's other parts that differ: {', '.join(f'{n} ({s})' for n, s in cmp.parts) if cmp.parts else 'none'}",
        '',
    ]
    pages = pdf_pages(pdf) if pdf else None
    for n, c in enumerate(cmp.changes, start=1):
        lines += [f'## {n}. {c.kind}: {c.where}', '']
        if pages is not None:
            cap = None
            m = re.search(r'\(([^)]*)\), row', c.where)
            if m:
                cap = m.group(1)
            found = page_of(c, pages, cap)
            page = f'- Page: {found[0]} of `{Path(pdf).name}` (where {found[1]} is)' if found else '- Page: not found in the PDF'
            lines += [page, '']
        for d in c.details:
            lines.append(f'- {d}')
        if c.details:
            lines.append('')
        if c.before:
            lines += ['Before:', '', '```text', c.before, '```', '']
        if c.after:
            lines += ['After:', '', '```text', c.after, '```', '']
    return '\n'.join(lines).rstrip() + '\n'


# ---------------------------------------------------------------------------
# A small correction in one place
# ---------------------------------------------------------------------------

_TEXT_NODE = re.compile(r'(<(?:w|m):t(?:\s[^>]*)?>)([^<]*)(</(?:w|m):t>)')
_UNESCAPE = {'&amp;': '&', '&lt;': '<', '&gt;': '>', '&quot;': '"', '&apos;': "'"}


def _unescape(t: str) -> str:
    return re.sub(r'&(?:amp|lt|gt|quot|apos);', lambda m: _UNESCAPE[m.group(0)], t)


def _escape(t: str) -> str:
    return t.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')


def _spans(xml: str, open_re: str, close: str, within: Tuple[int, int]) -> List[Tuple[int, int]]:
    """The spans of the elements opened by `open_re` and closed by `close` directly inside
    `within`, nested ones of the same kind counted as part of their outer one."""
    tokens = sorted([(m.start(), 1) for m in re.finditer(open_re, xml[within[0]:within[1]])] +
                    [(m.start(), -1) for m in re.finditer(re.escape(close), xml[within[0]:within[1]])])
    spans, depth, start = [], 0, 0
    for pos, step in tokens:
        if step == 1:
            if depth == 0:
                start = pos
            depth += 1
        else:
            depth -= 1
            if depth == 0:
                spans.append((within[0] + start, within[0] + pos + len(close)))
    return spans


def locate(xml: str, table: Optional[int] = None, row: Optional[int] = None, cell: Optional[int] = None,
           paragraph: Optional[str] = None) -> Tuple[int, int]:
    """The span, in the document's XML, of one table cell (its table, row and cell counted from 1)
    or of the one paragraph whose words hold `paragraph`."""
    body = (xml.index('<w:body>'), xml.index('</w:body>'))
    if table is not None:
        tables = _spans(xml, r'<w:tbl>', '</w:tbl>', body)
        if not 1 <= table <= len(tables):
            raise ValueError(f'the document has {len(tables)} tables, not {table}')
        rows = _spans(xml, r'<w:tr[\s>]', '</w:tr>', tables[table - 1])
        if row is None or not 1 <= row <= len(rows):
            raise ValueError(f'table {table} has {len(rows)} rows, not {row}')
        cells = _spans(xml, r'<w:tc[\s>]', '</w:tc>', rows[row - 1])
        if cell is None or not 1 <= cell <= len(cells):
            raise ValueError(f'row {row} of table {table} has {len(cells)} cells, not {cell}')
        return cells[cell - 1]
    if paragraph:
        found = [s for s in _spans(xml, r'<w:p[\s>]', '</w:p>', body)
                 if paragraph in ''.join(_unescape(m.group(2)) for m in _TEXT_NODE.finditer(xml[s[0]:s[1]]))]
        if len(found) != 1:
            raise ValueError(f'« {paragraph} » is in {len(found)} paragraphs, not exactly one')
        return found[0]
    raise ValueError('say where: a table, row and cell, or a paragraph')


def edit(src: Source, dest: Union[str, Path, BinaryIO], old: str, new: str, **where) -> None:
    """Replace `old` by `new` in one place of a Word document (`locate`'s table, row and cell, or
    paragraph), where `old` must be inside exactly one piece of text, exactly once. Every other
    part of the document is copied as it was, byte for byte."""
    with _zip(src) as z:
        infos = z.infolist()
        data = {i.filename: z.read(i.filename) for i in infos}
    xml = data[_DOCUMENT].decode('utf-8')
    start, end = locate(xml, **where)
    hits = [m for m in _TEXT_NODE.finditer(xml, start, end) if old in _unescape(m.group(2))]
    count = sum(_unescape(m.group(2)).count(old) for m in hits)
    if count != 1:
        raise ValueError(f'« {old} » is there {count} times, not exactly once'
                         + ('' if count else ' (a word split across two pieces of text cannot be replaced this way)'))
    m = hits[0]
    text = _escape(_unescape(m.group(2)).replace(old, new))
    data[_DOCUMENT] = (xml[:m.start(2)] + text + xml[m.end(2):]).encode('utf-8')
    with zipfile.ZipFile(dest, 'w') as out:
        for i in infos:
            out.writestr(i, data[i.filename], compress_type=i.compress_type)
