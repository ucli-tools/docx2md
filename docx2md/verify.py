"""
Verify a conversion: compare a Word document with the Markdown made from it.

Every check reads the .docx XML directly and the .md as text, so a defect
in the converter cannot hide itself. A check fails when something in the
document did not arrive in the Markdown (or arrived changed), and names
what, so the fix can go into the converter rather than into the output.
"""

import re
import zipfile
from collections import Counter
from dataclasses import dataclass, field
from pathlib import Path
from typing import Dict, List, Set
from xml.etree import ElementTree as ET

_W = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
_M = '{http://schemas.openxmlformats.org/officeDocument/2006/math}'
_A = '{http://schemas.openxmlformats.org/drawingml/2006/main}'

_RE_NUMBER = re.compile(r'#+\s*(?:\(\s*([0-9]+(?:\.[0-9]+)*[a-z]?)\s*\)|([0-9]+(?:\.[0-9]+)+[a-z]?))')
_RE_WORD = re.compile(r'[A-Za-z]{4,}')
_RE_CAPTION = re.compile(r'^\s*Figures?\s+[0-9]')


@dataclass
class Check:
    name: str
    ok: bool
    docx: str
    md: str
    details: List[str] = field(default_factory=list)


# ---------------------------------------------------------------------------
# The Word document
# ---------------------------------------------------------------------------

@dataclass
class _Docx:
    equations: int = 0
    numbers: Dict[str, str] = field(default_factory=dict)   # number -> punctuation
    images: int = 0
    captions: int = 0
    tables: int = 0
    headings: List[str] = field(default_factory=list)
    index_entries: int = 0
    footnotes: int = 0
    words: Counter = field(default_factory=Counter)
    text_letters: int = 0


def _read_docx(path: Path) -> _Docx:
    with zipfile.ZipFile(path) as z:
        root = ET.fromstring(z.read('word/document.xml'))
    parent = {c: p for p in root.iter() for c in p}
    d = _Docx()

    # Field state: instruction parts are hidden code; the results of TOC and
    # INDEX fields are Word's own generated text, rebuilt by the renderer
    hidden: Set[ET.Element] = set()
    generated: Set[ET.Element] = set()
    stack: List[Dict[str, str]] = []
    for el in root.iter():
        if el.tag == _W + 'fldChar':
            kind = el.get(_W + 'fldCharType')
            if kind == 'begin':
                stack.append({'state': 'code', 'code': ''})
            elif kind == 'separate' and stack:
                stack[-1]['state'] = 'result'
            elif kind == 'end' and stack:
                code = stack.pop()['code']
                if code.lstrip().startswith('XE') and re.search(r'XE\s+["\u201c]\s*\S', code):
                    d.index_entries += 1
        elif stack:
            top = stack[-1]
            if top['state'] == 'code':
                hidden.add(el)
                if el.tag == _W + 'instrText':
                    top['code'] += el.text or ''
                elif el.tag == _M + 'oMath':
                    top['code'] += 'M'
            elif any(f['state'] == 'result' and f['code'].split()[:1] in (['TOC'], ['INDEX'])
                     for f in stack):
                generated.add(el)

    def math_text(e: ET.Element) -> str:
        return ''.join(t.text or '' for t in e.iter(_M + 't'))

    # Equations and their numbers
    for e in root.iter(_M + 'oMathPara'):
        if e in hidden:
            continue
        d.equations += 1
    for e in root.iter(_M + 'oMath'):
        if e in hidden or parent.get(e) is not None and parent[e].tag == _M + 'oMathPara':
            continue
        text = math_text(e)
        # a lone accented letter typed as an equation becomes text, and an
        # equation with no math in it (only a footnote mark) is no equation
        if text.strip() and not (len(text) == 1 and not text.isascii() and text.isalpha()):
            d.equations += 1
    for e in list(root.iter(_M + 'oMathPara')) + [
            x for x in root.iter(_M + 'oMath') if parent[x].tag != _M + 'oMathPara']:
        if e in hidden:
            continue
        text = math_text(e)
        for m in _RE_NUMBER.finditer(text):
            before = text[:m.start()].rstrip()
            d.numbers[m.group(1) or m.group(2)] = before[-1] if before and before[-1] in '.,;:' else ''

    # Images, captions, tables
    # A caption is the paragraph right after a picture; a sentence that
    # begins "Figure 1.2.8 shows" is text, checked with the text
    d.images = sum(1 for _ in root.iter(_A + 'blip'))
    for holder in root.iter():
        previous = None
        for p in holder:
            if p.tag != _W + 'p':
                previous = p
                continue
            text = ''.join(t.text or '' for t in p.iter(_W + 't') if t not in hidden)
            has_image = p.find(f'.//{_A}blip') is not None
            if not text.strip() and not has_image:
                continue
            if (_RE_CAPTION.match(text) and previous is not None
                    and previous.find(f'.//{_A}blip') is not None):
                d.captions += 1
            previous = p
    for tbl in root.iter(_W + 'tbl'):
        nested = [t for t in tbl.iter(_W + 'tbl') if t is not tbl]
        outer = parent.get(tbl)
        while outer is not None and outer.tag != _W + 'tbl':
            outer = parent.get(outer)
        if outer is not None:
            continue                      # counted with its outer table
        d.tables += len(nested) if nested else 1

    # Headings (non-empty; math reduced to its text)
    for p in root.iter(_W + 'p'):
        style = p.find(f'{_W}pPr/{_W}pStyle')
        if style is None or not style.get(_W + 'val', '').startswith('Heading'):
            continue
        text = ''.join((t.text or '') for t in p.iter(_W + 't') if t not in hidden)
        if text.strip():
            d.headings.append(text)

    # Text words, paragraph by paragraph (Word splits words across runs)
    for p in root.iter(_W + 'p'):
        if p in generated or p in hidden:
            continue
        text = ''.join(t.text or '' for t in p.iter(_W + 't')
                       if t not in hidden and t not in generated)
        d.words.update(_RE_WORD.findall(text))

    d.footnotes = sum(1 for e in root.iter()
                      if e.tag in (_W + 'footnoteReference', _W + 'endnoteReference')
                      and e not in hidden)
    return d


# ---------------------------------------------------------------------------
# The Markdown
# ---------------------------------------------------------------------------

@dataclass
class _Md:
    equations: int = 0
    numbers: Dict[str, str] = field(default_factory=dict)
    images: int = 0
    captions: int = 0
    lost_captions: List[str] = field(default_factory=list)
    tables: int = 0
    headings: List[str] = field(default_factory=list)
    index_markers: int = 0
    footnotes: int = 0
    unmatched_notes: List[str] = field(default_factory=list)
    words: Counter = field(default_factory=Counter)
    leftovers: List[str] = field(default_factory=list)


_RE_DISPLAY = re.compile(r'^\$\$\n(.*?)\n\$\$(?:\[\^[^\]\s]+\])*$', re.S | re.M)
_RE_RAW_DISPLAY = re.compile(r'```\{=latex\}\n\\\[(.*?)\\\]\n```', re.S)
_RE_INLINE = re.compile(r'(?<![\\$])\$(?!\$)((?:\\\$|[^$])+?)\$')
_RE_TAG = re.compile(r'\\tag\*?\{([^}]*)\}')


def _ending_punctuation(latex: str) -> str:
    body = _RE_TAG.split(latex)[0]
    # look through closing groups: Word may keep the mark inside one, as in
    # m_{2.} or \ln{(x),}, and renders it where LaTeX does
    body = re.sub(r'(\s|\\end\{[a-z*]+\}|\\\\|\}|\\right\.|\\\s)+$', '', body)
    body = body.rstrip()
    return body[-1] if body and body[-1] in '.,;:' else ''


def _count_tables(markdown: str) -> int:
    """Tables as pandoc reads them: pipe, grid, simple and multiline alike."""
    import json
    import pypandoc

    ast = json.loads(pypandoc.convert_text(markdown, 'json', format='markdown'))
    count = 0
    stack = [ast['blocks']]
    while stack:
        node = stack.pop()
        if isinstance(node, dict):
            if node.get('t') == 'Table':
                count += 1
            stack.extend(node.values())
        elif isinstance(node, list):
            stack.extend(node)
    return count


def _detached_scripts(text: str) -> List[str]:
    """Scripts set on an empty group ``{}`` right after another group.

    Only a second script of the same kind needs that (``x_{1}{}_{2}``). After
    a subscript, ``_{k=0}{}^{n}`` moves a limit off its sign; after any other
    group, ``\\mathbb{C}{}^{3}`` moves the power off its base.
    """
    def opening(close: int) -> int:
        depth, j = 0, close
        while j >= 0:
            if text[j] == '}' and text[j - 1] != '\\':
                depth += 1
            elif text[j] == '{' and text[j - 1] != '\\':
                depth -= 1
                if depth == 0:
                    return j
            j -= 1
        return -1

    found = []
    for m in re.finditer(r'\}\{\}([_^])', text):
        close = m.start()
        # A double script, direct (x_{1}{}_{2}) or through braces TeX sees
        # through ({{X}_{1}}{}_{1,j}): the empty group is needed there
        needed = False
        while close >= 0 and text[close] == '}':
            j = opening(close)
            if j > 0 and text[j - 1] == m.group(1):
                needed = True
                break
            close -= 1
        if not needed:
            j = opening(m.start())
            found.append(text[max(j - 12, 0):m.end() + 6].replace('\n', ' '))
    return found


def _read_md(path: Path) -> _Md:
    text = path.read_text(encoding='utf-8')
    if text.startswith('---\n'):
        end = text.find('\n---\n', 4)
        text = text[end + 5:] if end != -1 else text
    m = _Md()

    displays = _RE_DISPLAY.findall(text)
    for block in _RE_RAW_DISPLAY.findall(text):
        # a wide equation: the math sits in \sbox0{$\displaystyle ...$},
        # its number after the box
        boxed = re.search(r'\\sbox0\{\$\\displaystyle\s*(.*?)\$\}%', block, re.S)
        tags = ''.join(f' \\tag{{{t}}}' for t in _RE_TAG.findall(block))
        displays.append((boxed.group(1) if boxed else block) + tags)
    rest = _RE_RAW_DISPLAY.sub(' ', _RE_DISPLAY.sub(' ', text))
    rest_no_index = re.sub(r'\[index:[^\]]*\]', ' ', rest)
    inlines = _RE_INLINE.findall(rest_no_index)
    m.equations = len(displays) + len(inlines)
    for latex in displays + inlines:
        for number in _RE_TAG.findall(latex):
            m.numbers[number] = _ending_punctuation(latex)

    blocks = re.split(r'\n\s*\n', text)
    for k, block in enumerate(blocks):
        images = re.findall(r'!\[((?:[^\[\]]|\[[^\]]*\])*)\]\(', block)
        if not images:
            continue
        m.images += len(images)
        if len(images) == 1:
            if _RE_CAPTION.match(images[0]):
                m.captions += 1
        else:
            after = blocks[k + 1].strip() if k + 1 < len(blocks) else ''
            if _RE_CAPTION.match(after.lstrip('*_ ')):
                m.captions += 1
            elif any(_RE_CAPTION.match(a) for a in images):
                m.lost_captions.append(images[0][:70])

    m.tables = _count_tables(text)

    m.headings = [re.sub(r'^#+\s*', '', h) for h in re.findall(r'^#{1,6} .*$', text, re.M)]
    m.index_markers = len(re.findall(r'\[index:[^\]]*\]', text))
    defined = re.findall(r'^\[\^([^\]]+)\]:', text, re.M)
    used = [m.group(1) for m in re.finditer(r'\[\^([^\]\s]+)\]', text)
            if not (text[m.end():m.end() + 1] == ':'
                    and (m.start() == 0 or text[m.start() - 1] == '\n'))]
    m.footnotes = len(used)
    m.unmatched_notes = sorted(set(used) ^ set(defined))

    prose = _RE_INLINE.sub(' ', rest)
    prose = re.sub(r'\[index:[^\]]*\]', ' ', prose)
    prose = re.sub(r'```.*?```', ' ', prose, flags=re.S)
    prose = re.sub(r'\]\([^)]*\)', ']', prose)      # keep captions, drop image paths
    prose = re.sub(r'<https?://[^>]*>', lambda u: ' ' + u.group(0)[1:-1] + ' ', prose)
    m.words.update(_RE_WORD.findall(prose))

    if '@@' in text:
        m.leftovers.append(f"{text.count('@@') // 2} converter placeholders (@@...@@)")
    if re.search(r'\\?#\(\s*[0-9]+\.[0-9]', text):
        m.leftovers.append('Word equation numbers left as #(n)')
    stray = [c for c in _RE_INLINE.sub(' ', rest) if 0x1D400 <= ord(c) <= 0x1D7FF]
    if stray:
        m.leftovers.append(f'{len(stray)} math letters in running text ({"".join(sorted(set(stray)))[:10]})')
    detached = _detached_scripts(text)
    if detached:
        m.leftovers.append(f'{len(detached)} scripts moved off their base by an empty group ({detached[0]})')
    split = re.findall(r'[^\W\d_]+\[index:[^\]]*\][^\W\d_]+', text)
    if split:
        m.leftovers.append(f'{len(split)} words split by an index marker ({split[0][:40]})')
    controls = [c for c in text if ord(c) < 32 and c not in '\n\t']
    if controls:
        m.leftovers.append(f'{len(controls)} control characters')
    return m


# ---------------------------------------------------------------------------
# Comparison
# ---------------------------------------------------------------------------

def _norm_heading(h: str) -> str:
    h = re.sub(r'\$[^$]*\$', ' ', h)
    h = re.sub(r'\[index:[^\]]*\]', ' ', h)
    return re.sub(r'[^a-z0-9]', '', h.lower())


def verify(docx_path, md_path, word_tolerance: float = 0.005) -> List[Check]:
    """Compare *docx_path* with the Markdown made from it."""
    d = _read_docx(Path(docx_path))
    m = _read_md(Path(md_path))
    checks: List[Check] = []

    # Every equation is in the Markdown: allow the converter's own text-mode
    # additions (math letters typed in text become math), never a loss
    missing_eqs = d.equations - m.equations
    checks.append(Check(
        'equations', missing_eqs <= 0, str(d.equations), str(m.equations),
        [] if missing_eqs <= 0 else [f'{missing_eqs} equations of the .docx are not in the .md']))

    lost = sorted(set(d.numbers) - set(m.numbers))
    extra = sorted(set(m.numbers) - set(d.numbers))
    checks.append(Check(
        'equation numbers', not lost and not extra, str(len(d.numbers)), str(len(m.numbers)),
        ([f'missing: {", ".join(lost[:12])}'] if lost else [])
        + ([f'not in the .docx: {", ".join(extra[:12])}'] if extra else [])))

    wrong = [(n, p, m.numbers[n]) for n, p in d.numbers.items()
             if n in m.numbers and m.numbers[n] != p]
    checks.append(Check(
        'punctuation before equation numbers', not wrong,
        f'{sum(1 for p in d.numbers.values() if p)} with punctuation',
        f'{len(wrong)} differ',
        [f'({n}): .docx "{p}" .md "{q}"' for n, p, q in wrong[:10]]))

    checks.append(Check('images', d.images == m.images, str(d.images), str(m.images)))
    checks.append(Check(
        'figure captions', m.captions == d.captions and not m.lost_captions,
        str(d.captions), str(m.captions),
        [f'caption not shown for side-by-side images: {c}' for c in m.lost_captions]))
    checks.append(Check('tables', m.tables >= d.tables, str(d.tables), str(m.tables)))

    md_heads = Counter(_norm_heading(h) for h in m.headings)
    missing_heads = [h for h in d.headings
                     if _norm_heading(h) and _norm_heading(h) not in md_heads
                     and not (h.strip().lower() == 'index' and m.index_markers)]
    checks.append(Check(
        'headings', not missing_heads, str(len(d.headings)), str(len(m.headings)),
        [f'missing: {h[:70]}' for h in missing_heads[:10]]))

    checks.append(Check(
        'index entries', d.index_entries == m.index_markers or d.index_entries == 0,
        str(d.index_entries), str(m.index_markers)))
    checks.append(Check(
        'footnotes', m.footnotes == d.footnotes and not m.unmatched_notes,
        str(d.footnotes), str(m.footnotes),
        [f'note without its text or mark: [^{n}]' for n in m.unmatched_notes]))

    total = sum(d.words.values())
    gone = d.words - m.words
    lost_words = sum(gone.values())
    checks.append(Check(
        'text', total == 0 or lost_words / total <= word_tolerance,
        f'{total} words', f'{lost_words} missing ({lost_words / max(total, 1):.2%})',
        [f'most missing: {", ".join(f"{w} ({n})" for w, n in gone.most_common(8))}'] if lost_words else []))

    checks.append(Check('leftovers', not m.leftovers, '-', str(len(m.leftovers)), m.leftovers))
    return checks
