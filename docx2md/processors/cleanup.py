"""
Word document cleanup processor for docx2md.

Removes Word-specific artifacts from raw pandoc markdown output:
- Duplicate OMML+text equation representations ($X$$Y$$$  →  $X$)
- Auto-generated Table of Contents section
- Bold/italic markup inside headings (# ***Title*** → # Title)
- Word-generated heading IDs ({#word-id}), keeping .unnumbered
- Image size attributes ({width="..." height="..."})
- Absolute image paths → relative paths
- Structural LaTeX validation (brace balance in captions, double subscripts)
"""

import os
import re
from pathlib import Path
from typing import Any, Dict, Optional

from docx2md.processors.base import BaseProcessor
from docx2md.utils.logging_utils import get_logger

logger = get_logger(__name__)

# ---------------------------------------------------------------------------
# Triple-dollar ($$$) artifact patterns
# ---------------------------------------------------------------------------
# Word documents often store equations BOTH as OMML (→ LaTeX by pandoc) AND as
# a plain-text fallback label.  Pandoc converts both, producing e.g.:
#   $S_{3}$$groups:$$$   =  $S_{3}$  +  $groups:$  +  $$ (orphan display open)
#   $Z_{3}$$Z_{3}$$$     =  $S_{3}$  +  $Z_{3}$  (duplicate) +  $$ (orphan)
# Both end with exactly three consecutive $ signs.  LaTeX errors result when
# pandoc re-encodes these as \[ ... \] with $ inside.
#
# Pattern A – two inline-math spans followed by orphan $$
#   $X$$Y$$$  →  $X$   (keep LaTeX version, drop text label + orphan opener)
_P_TRIPLE_DOLLAR_A = re.compile(
    r'(\$[^$\n]+\$)'    # group 1: first inline math (LaTeX-formatted)
    r'\$[^$\n]+\$'      # second inline math (text label – discard)
    r'\$\$(?!\$)',       # orphan display opener $$ (not $$$$)
)

# Pattern B – stray $ immediately before a display math span
#   $$$X$$  →  $$X$$
_P_TRIPLE_DOLLAR_B = re.compile(
    r'(?<!\$)\$(?=\$\$(?!\$))',  # lone $ right before $$...  but not $$$$
)

# Pattern C – five+ consecutive $ collapsed to $$ (edge case triple/quad overlap)
_P_MANY_DOLLARS = re.compile(r'\${5,}')
_P_FOUR_DOLLARS = re.compile(r'(?<!\$)\$\$\$\$(?!\$)')

# Pattern D – $$WORD$$ text label followed immediately by raw LaTeX commands.
# Word stores variable definitions like "for unitary U ∈ ℂ^{m×m}" as two equations:
# a text label ("Unitary") and the actual math ("U ∈ ℂ^{m×m}").  Pandoc produces:
#   for unitary$$Unitary$$\ U \in \mathbb{C}^{m \times m}
# The $$Unitary$$ becomes a display-math block (\[Unitary\]) followed by raw LaTeX
# in text mode, causing "Missing $ inserted."  Fix: remove the $$WORD$$ text label
# and wrap the raw-LaTeX tail in inline math $...$:
#   for unitary $\ U \in \mathbb{C}^{m \times m}$
# Matches: anything on the line, then $$WORD$$, then tail starting with backslash.
_P_WORD_LABEL_INLINE = re.compile(
    r'(?m)'                      # multiline: ^ $ match line start/end
    r'^(.+?)'                    # group 1: text before $$WORD$$
    r'\$\$([A-Za-z]+)\$\$'       # $$WORD$$ text label (alphabetic only)
    r'(\\[^$\n]+)'               # group 3: tail starting with \ and containing NO $
    r'\$?$'                      # optional orphan $ before end of line (discard it)
)

# Pattern E – double subscript from Word's nested equation encoding.
# Word produces e.g. {{\widehat{X}}_{1}}_{1,j} — two subscripts on the same
# base.  LaTeX errors: "! Double subscript."  Fix: insert {} between them so
# the second subscript attaches to an empty group rather than re-subscribing:
#   }_{1}}_{1,j}  →  }_{1}}{}_{1,j}
_P_DOUBLE_SUBSCRIPT = re.compile(
    r'(\}_{[^}]*\})'   # group 1: subscript closing, e.g. }_{1}}
    r'\}'               # closing brace of the outer group
    r'(_{)',            # group 2: second subscript opener _{
)

# Pattern F – Word underline spans: [text]{.underline} → text
# Pandoc converts Word underline formatting to [text]{.underline}, which
# becomes \ul{text} in LaTeX — undefined without the soul package.
# Strip the span and keep the text.
_P_UNDERLINE_SPAN = re.compile(r'\[([^\]]*)\]\{\.underline\}')

# ---------------------------------------------------------------------------
# Structural LaTeX sanitizers (class-level, not pattern-specific)
# ---------------------------------------------------------------------------

# Pattern G – image alt text: ![ALT](path)
# Pandoc puts alt text into \caption{}.  Captions break on:
#   - Display math $$...$$ (illegal inside \caption)
#   - Unbalanced braces (unclosed { from garbled Word math)
# Fix: convert $$ to $ in alt text, then strip any math span with
# unbalanced braces.
_P_IMAGE_LINE = re.compile(r'!\[([^\]]*)\]\(([^)]+)\)')

# Pattern H – pandoc superscript ^text^ glued to inline math $...$
# Word footnotes/references produce:  $\theta$^n^ = 1
# Pandoc makes this: inline math + \textsuperscript{n} + text " = 1"
# LaTeX chokes because ^n^ outside math mode is ambiguous.
# Fix: merge $X$^Y^ into $X^{Y}$ (single math span with superscript).
_P_MATH_THEN_SUPER = re.compile(
    r'\$([^$]+)\$'            # group 1: inline math content
    r'\^([^^\n{}'r']+)\^',   # group 2: pandoc superscript (no braces)
)

# Pattern I – missing space after closing inline math: $...$word → $...$ word
# Pandoc sometimes glues closing $ to the next word when Word had no separator.
# The [^ $\n] after opening $ ensures we match real inline math (no space after $)
# and don't span from one closing $ to the next opening $ across normal text.
_P_MATH_GLUED_AFTER = re.compile(
    r'(\$[^ $\n][^$\n]*\$)([a-zA-Z])',
)

# Pattern J – Word text label glued to inline math of the same word:
#   "unitary$Unitary$" → "unitary" (remove duplicate math label)
# Word stores both plain text and an equation label for the same word.
_P_WORD_TEXT_LABEL = re.compile(
    r'([a-zA-Z]{2,})\$([a-zA-Z]+)\$',
)

# Pattern K – broken inline math with space after opening $: $ n → $n
# Pandoc requires no space after opening $ for inline math recognition.
# Various pipeline interactions (math extraction _fix_adjacent_inline,
# pattern D rewrites) can produce $ followed by a space.  Fix by removing
# the space.  Works for both single-line ($n$) and multiline (matrices).
# Guard: negative lookbehind excludes CLOSING $ (preceded by math content:
# letters, digits, }, ), ], \) and $$ display math (preceded by $).
# Opening $ is preceded by whitespace, formatting chars (*_), punctuation,
# or start-of-line — none of which are in the exclusion set.
_P_SPACE_AFTER_OPEN_DOLLAR = re.compile(
    r'(?<![a-zA-Z0-9})\]\\$])\$ ([a-zA-Z\\])',
)

# ---------------------------------------------------------------------------
# Matches Word-generated TOC headings (with optional class attributes)
_TOC_HEADING = re.compile(
    r'^#+\s+\*{0,3}Table\s+of\s+Contents\*{0,3}(?:\s*\{[^}]*\})?\s*$',
    re.IGNORECASE,
)

# Matches heading lines (one or more # followed by space and content)
_HEADING_LINE = re.compile(r'^(#+)\s+(.+)$')

# Empty headings: lines with only # characters and optional whitespace
_EMPTY_HEADING = re.compile(r'^#+\s*$', re.MULTILINE)

# Empty bracket artifacts: standalone [] on a line (from empty Word links/images)
_EMPTY_BRACKETS = re.compile(r'^\[\]\s*$', re.MULTILINE)

# Matches {#word-id} optionally followed by class attributes like .unnumbered
_HEADING_ID = re.compile(r'\{#[a-zA-Z0-9_-]+([^}]*)\}')

# Image with size attributes: ![alt](path){width="..." height="..."}
_IMAGE_SIZE_ATTRS = re.compile(
    r'(!\[[^\]]*\]\([^)]+\))\{(?:width|height)="[^"]*"(?:\s+(?:width|height)="[^"]*")?\}',
)

# Image markdown reference: ![alt](path)
_IMAGE_REF = re.compile(r'!\[([^\]]*)\]\(([^)]+)\)')


class WordCleanupProcessor(BaseProcessor):
    """
    Cleans up Word-specific artifacts in pandoc-generated markdown.

    Args:
        config: Configuration dict (uses config["cleanup"] section)
        output_dir: Directory of the output file — used to make image paths relative
    """

    def __init__(
        self,
        config: Optional[Dict[str, Any]] = None,
        output_dir: Optional[Path] = None,
    ):
        super().__init__(config)
        self.output_dir = output_dir or Path('.')
        cfg = (config or {}).get('cleanup', {})
        self.strip_triple_dollar = cfg.get('strip_triple_dollar', True)
        self.remove_toc = cfg.get('remove_toc', True)
        self.strip_heading_markup = cfg.get('strip_heading_markup', True)
        self.strip_heading_ids = cfg.get('strip_heading_ids', True)
        self.remove_image_attrs = cfg.get('remove_image_attrs', True)
        self.fix_image_paths = cfg.get('fix_image_paths', True)

    def process(self, content: str) -> str:
        # Must run first – $$$  corrupts math tokenizers downstream
        if self.strip_triple_dollar:
            content = _strip_triple_dollar(content)
        if self.remove_toc:
            content = _remove_toc(content)
        if self.strip_heading_markup:
            content = _strip_heading_markup(content)
        if self.strip_heading_ids:
            content = _strip_heading_ids(content)
        if self.remove_image_attrs:
            content = _IMAGE_SIZE_ATTRS.sub(r'\1', content)
        if self.fix_image_paths:
            content = _fix_image_paths(content, self.output_dir)
        # Strip empty headings (## with no text) — Word section break artifacts
        content = _EMPTY_HEADING.sub('', content)
        # Strip empty bracket artifacts [] on standalone lines
        content = _EMPTY_BRACKETS.sub('', content)
        return content


def _remove_toc(content: str) -> str:
    """
    Remove the Word auto-generated Table of Contents section.

    Drops the TOC heading and all following TOC link lines until the next
    real heading (a heading line that doesn't contain a pandoc TOC link).
    """
    lines = content.split('\n')
    result = []
    in_toc = False

    for line in lines:
        if _TOC_HEADING.match(line):
            in_toc = True
            continue  # drop TOC heading itself

        if in_toc:
            # A real heading that is NOT a TOC hyperlink ends the block
            if _HEADING_LINE.match(line) and '](#' not in line:
                in_toc = False
                result.append(line)
            # Otherwise silently drop TOC content (links, blank lines)
        else:
            result.append(line)

    return '\n'.join(result)


def _strip_heading_markup(content: str) -> str:
    """
    Strip bold/italic markers (*** ** *) from heading text.

    Handles mixed headings like:
        # ***Title with bold*** $math$
        ## **Partial bold** remaining text
    """
    lines = content.split('\n')
    result = []
    for line in lines:
        m = _HEADING_LINE.match(line)
        if m:
            hashes = m.group(1)
            text = m.group(2)
            # Iteratively strip leading/trailing *** ** * from first text segment
            # Stop when text doesn't change (avoids infinite loop on nested markers)
            prev = None
            while prev != text:
                prev = text
                # Strip *** bold-italic wrapping the first segment (may be followed by math)
                text = re.sub(r'^\*{1,3}(.*?)\*{1,3}(\s)', r'\1\2', text)
                text = re.sub(r'^\*{1,3}(.*?)\*{1,3}$', r'\1', text)
            result.append(f'{hashes} {text}')
        else:
            result.append(line)
    return '\n'.join(result)


def _strip_heading_ids(content: str) -> str:
    """
    Remove Word-generated heading IDs but keep pandoc's .unnumbered class.

    {#word-id}             → (removed)
    {#word-id .unnumbered} → {.unnumbered}
    {#id .TOC-Heading .unnumbered} → {.unnumbered}
    """
    def _replace(match: re.Match) -> str:
        extra = match.group(1).strip()
        if '.unnumbered' in extra:
            return ' {.unnumbered}'
        return ''

    return _HEADING_ID.sub(_replace, content)


def _fix_image_paths(content: str, output_dir: Path) -> str:
    """
    Convert absolute image paths to paths relative to output_dir.
    """
    def _replace(match: re.Match) -> str:
        alt = match.group(1)
        path_str = match.group(2)
        p = Path(path_str)
        if p.is_absolute():
            try:
                rel = os.path.relpath(p, output_dir)
                return f'![{alt}]({rel})'
            except ValueError:
                pass  # different drive on Windows — leave as-is
        return match.group(0)

    return _IMAGE_REF.sub(_replace, content)


def _strip_triple_dollar(content: str) -> str:
    """
    Remove duplicate OMML+text equation pairs that pandoc emits as $$$...$$$.

    Word stores equations twice: once as OMML (which pandoc renders as proper
    LaTeX inline math $...$) and once as a plain-text label (which pandoc also
    renders as inline math $...$).  The label is then followed by an orphaned
    display-math opener $$, producing invalid sequences like:

        $S_{3}$$groups:$$$   →  $S_{3}$
        $Z_{3}$$Z_{3}$$$     →  $Z_{3}$
        $$$Unitary$$         →  $$Unitary$$

    pdflatex errors on these because $ is illegal inside \\[...\\] display mode.
    """
    # Edge case: 5+ consecutive $ → reduce to $$
    content = _P_MANY_DOLLARS.sub('$$', content)

    # Edge case: exactly 4 consecutive $ (adjacent display math) → $$
    content = _P_FOUR_DOLLARS.sub('$$', content)

    # Pattern A: $X$$Y$$$  →  $X$
    # Keep the first (LaTeX) inline math, drop text label and orphan $$
    content = _P_TRIPLE_DOLLAR_A.sub(r'\1', content)

    # Pattern B: $$$X$$  →  $$X$$
    # Remove stray $ immediately before a display math span
    content = _P_TRIPLE_DOLLAR_B.sub('', content)

    # Pattern D: TEXT$$WORD$$\LATEX...  →  TEXT $\LATEX...$
    # Remove inline text-label $$WORD$$ and wrap the raw-LaTeX tail in inline math.
    content = _P_WORD_LABEL_INLINE.sub(
        lambda m: f'{m.group(1)} ${m.group(3)}$', content
    )

    # Pattern E: fix double subscripts }_{X}}_{Y} → }_{X}}{}_{Y}
    content = _P_DOUBLE_SUBSCRIPT.sub(r'\1}{}\2', content)

    # Pattern F: strip Word underline spans [text]{.underline} → text
    content = _P_UNDERLINE_SPAN.sub(r'\1', content)

    # Pattern G: sanitize image alt text (display math, unbalanced braces)
    content = _sanitize_image_alt(content)

    # Pattern H: merge $X$^Y^ → $X^{Y}$ (pandoc superscript after inline math)
    content = _P_MATH_THEN_SUPER.sub(r'$\1^{\2}$', content)

    # Pattern J: word$Word$ duplicate text label → word (remove math duplicate)
    content = _P_WORD_TEXT_LABEL.sub(
        lambda m: m.group(1) if m.group(1).lower() == m.group(2).lower()
        else m.group(0),
        content,
    )

    # Pattern I: $...$word → $...$ word (ensure space after closing inline math)
    content = _P_MATH_GLUED_AFTER.sub(r'\1 \2', content)

    # Pattern K: $ n$ → $n$ (fix space after opening $ that breaks inline math)
    content = _P_SPACE_AFTER_OPEN_DOLLAR.sub(r'$\1', content)

    return content


def _sanitize_image_alt(content: str) -> str:
    """
    Fix LaTeX-breaking patterns inside image alt text (which becomes \\caption{}).

    1. Convert display math $$ to inline math $ inside alt text
    2. Strip any remaining math span that has unbalanced braces
    """
    def _fix_alt(m: re.Match) -> str:
        alt = m.group(1)
        path = m.group(2)

        # Convert \[...\] and \(...\) to $...$ first
        alt = re.sub(r'\\\[', '$', alt)
        alt = re.sub(r'\\\]', '$', alt)
        alt = re.sub(r'\\\(', '$', alt)
        alt = re.sub(r'\\\)', '$', alt)
        # Convert $$ → $ (display math illegal in captions)
        alt = alt.replace('$$', '$')

        # Check brace balance in each math span; strip broken ones
        result = []
        i = 0
        while i < len(alt):
            if alt[i] == '$' and (i == 0 or alt[i - 1] != '\\'):
                # Find closing $
                j = alt.find('$', i + 1)
                if j == -1:
                    # No closing $ — strip from here to end
                    break
                span = alt[i + 1:j]
                if span.count('{') == span.count('}'):
                    result.append(f'${span}$')
                # else: skip this broken math span
                i = j + 1
            else:
                result.append(alt[i])
                i += 1

        alt_fixed = ''.join(result)
        return f'![{alt_fixed}]({path})'

    return _P_IMAGE_LINE.sub(_fix_alt, content)


# Bare LaTeX command followed by pandoc superscript: \theta^n^ → $\theta^{n}$
# Pandoc superscripts are short (1-20 chars), never cross newlines, and never
# contain braces {} (which would indicate LaTeX ^{...} syntax, not pandoc).
_P_BARE_CMD_SUPER = re.compile(
    r'(?<!\$)'                # not already inside math (no preceding $)
    r'(\\[a-zA-Z]+)'         # group 1: LaTeX command like \theta
    r'\^([^^\n{}'r']{1,20})\^'  # group 2: pandoc superscript (no braces/^/newline)
    r'(?!\$)',                # not followed by $ (not already in math)
)


def final_sanitize(content: str) -> str:
    """
    Final-pass LaTeX sanitization.

    Runs AFTER fix_delimiters (which may introduce new $$ or change delimiter
    forms) and AFTER FigureProcessor (which may set new alt text).  Catches
    issues that the early cleanup pass couldn't see.
    """
    # NOTE: _strip_dollars_in_display_math disabled — it over-aggressively strips
    # $ signs inside $$...$$ blocks, breaking valid LaTeX constructs.  The root
    # causes (e.g. _P_BARE_CMD_SUPER matching LaTeX ^{} syntax) are now fixed
    # with brace-exclusion in the regexes.
    # content = _strip_dollars_in_display_math(content)

    # Pattern E: double subscripts (re-run in case fix_delimiters created new ones)
    content = _P_DOUBLE_SUBSCRIPT.sub(r'\1}{}\2', content)

    # Pattern F: underline spans
    content = _P_UNDERLINE_SPAN.sub(r'\1', content)

    # Pattern G: sanitize image alt text (display math, unbalanced braces)
    content = _sanitize_image_alt(content)

    # Pattern H: merge $X$^Y^ → $X^{Y}$ (pandoc superscript after inline math)
    content = _P_MATH_THEN_SUPER.sub(r'$\1^{\2}$', content)

    # Bare LaTeX command + pandoc superscript: \theta^n^ → $\theta^{n}$
    content = _P_BARE_CMD_SUPER.sub(r'$\1^{\2}$', content)

    # Pattern K: $ n$ → $n$ (fix space after opening $ — final safety net)
    content = _P_SPACE_AFTER_OPEN_DOLLAR.sub(r'$\1', content)

    return content


def _strip_dollars_in_display_math(content: str) -> str:
    """
    Remove lone $ signs inside $$...$$ display math blocks.

    Word text labels leak into display math, producing e.g.:
        $$... = $\\Sigma^{{\\dagger}U}${\\dagger}.$$
    The $...$ is a duplicate text label.  Removing the lone $ signs gives:
        $$... = \\Sigma^{{\\dagger}U}{\\dagger}.$$
    which compiles (if not perfectly formatted).
    """
    # Match $$...$$ blocks (possibly multi-line via \n in $$...$$)
    def _clean_block(m: re.Match) -> str:
        inner = m.group(1)
        # Remove lone $ that are not part of $$ (already stripped by outer match)
        cleaned = inner.replace('$', '')
        return f'$${cleaned}$$'

    # Match display math: $$ followed by content, then $$
    # Use non-greedy to avoid spanning multiple display blocks
    return re.sub(r'\$\$(.*?)\$\$', _clean_block, content, flags=re.DOTALL)
