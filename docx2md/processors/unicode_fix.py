"""
Unicode-to-LaTeX replacement processor for docx2md.

Replaces Unicode math symbols that break pdflatex with their LaTeX equivalents.
Applies context-aware replacements: different substitutions inside vs. outside
math delimiters ($...$ and $$...$$).
"""

import re
from typing import Any, Dict, List, Optional, Tuple

from docx2md.processors.base import BaseProcessor
from docx2md.utils.logging_utils import get_logger
from docx2md.utils.math_utils import tokenize_math_spans, reassemble

logger = get_logger(__name__)

# Replacements that are always safe regardless of math/text context
_ALWAYS: List[Tuple[str, str]] = [
    ('\u00a0', ' '),    # Non-breaking space
    ('≔',      ':='),   # U+2254 COLON EQUALS (used in math defs in text)
    ('Τ',      'T'),    # U+03A4 GREEK CAPITAL TAU → Latin T (same glyph, pdflatex-safe)
]

# Replacements inside math context ($...$ or $$...$$)
_IN_MATH: List[Tuple[str, str]] = [
    ('ℓ',  r'\ell '),   # U+2113 SCRIPT SMALL L
    ('Β',  'B'),         # U+0392 GREEK CAPITAL BETA (same glyph as italic B)
    ('θ',  r'\theta'),   # U+03B8 GREEK SMALL THETA
    ('⋰',  r'\ddots'),   # U+22F0 UP RIGHT DIAGONAL ELLIPSIS (safe fallback)
    ('₀',  '_0'), ('₁', '_1'), ('₂', '_2'), ('₃', '_3'), ('₄', '_4'),
    ('₅',  '_5'), ('₆', '_6'), ('₇', '_7'), ('₈', '_8'), ('₉', '_9'),
]

# Replacements outside math context (body text)
# Subscript digits in text become inline math
_IN_TEXT: List[Tuple[str, str]] = [
    ('Β',  'B'),
    ('θ',  '$\\theta$'),  # U+03B8 GREEK SMALL THETA → inline math in text
]

# Combined ell+subscript digit sequence in text: ℓ₁ → $\ell_1$
_ELL_SUB_TEXT = re.compile(r'ℓ([₀₁₂₃₄₅₆₇₈₉]+)')
_SUB_DIGIT_MAP = str.maketrans('₀₁₂₃₄₅₆₇₈₉', '0123456789')

# Lone ell in text (not followed by subscript digit)
_ELL_TEXT = re.compile(r'ℓ')

# Any remaining subscript digit in text
_SUB_DIGIT_TEXT = re.compile(r'([₀₁₂₃₄₅₆₇₈₉]+)')


class UnicodeFixProcessor(BaseProcessor):
    """
    Replaces Unicode math characters with LaTeX-safe equivalents.

    Uses math tokenization to apply different replacements inside vs. outside
    math delimiters, preventing double-escaping.

    Args:
        config: Configuration dict (uses config["unicode_fix"] section)
    """

    def __init__(self, config: Optional[Dict[str, Any]] = None):
        super().__init__(config)
        cfg = (config or {}).get('unicode_fix', {})
        self.enabled = cfg.get('enabled', True)
        self.custom = cfg.get('custom_replacements', [])

    def process(self, content: str) -> str:
        if not self.enabled:
            return content

        # Step 1: context-free replacements (safe anywhere)
        for old, new in _ALWAYS:
            content = content.replace(old, new)

        # Apply custom always-replacements from config
        for rule in self.custom:
            if 'char' in rule and 'always' in rule:
                content = content.replace(rule['char'], rule['always'])

        # Step 2: tokenize and apply context-aware replacements
        tokens = tokenize_math_spans(content)
        result = []
        for kind, text in tokens:
            if kind == 'math':
                text = _fix_in_math(text, self.custom)
            else:
                text = _fix_in_text(text, self.custom)
            result.append(text)

        return ''.join(result)


def _fix_in_math(text: str, custom: list) -> str:
    for old, new in _IN_MATH:
        text = text.replace(old, new)
    for rule in custom:
        if 'char' in rule and 'math' in rule:
            text = text.replace(rule['char'], rule['math'])
    return text


# Mathematical alphanumeric symbols (U+1D400-U+1D7FF) typed in body text:
# no text font has them, so each run becomes one inline equation
_MATH_ALNUM_RUN = re.compile('[\U0001D400-\U0001D7FF]+')
_MATH_ALNUM_STYLES = (
    ('MATHEMATICAL BOLD ITALIC ', '\\boldsymbol{%s}'),
    ('MATHEMATICAL BOLD ', '\\mathbf{%s}'),
    ('MATHEMATICAL ITALIC ', '%s'),
    ('MATHEMATICAL DOUBLE-STRUCK ', '\\mathbb{%s}'),
    ('MATHEMATICAL SCRIPT ', '\\mathcal{%s}'),
    ('MATHEMATICAL FRAKTUR ', '\\mathfrak{%s}'),
    ('MATHEMATICAL SANS-SERIF ', '\\mathsf{%s}'),
    ('MATHEMATICAL MONOSPACE ', '\\mathtt{%s}'),
)


def _math_alnum_latex(ch: str) -> str:
    import unicodedata
    name = unicodedata.name(ch, '')
    for prefix, template in _MATH_ALNUM_STYLES:
        if name.startswith(prefix):
            rest = name[len(prefix):]
            if rest.startswith('CAPITAL '):
                letter = rest[len('CAPITAL '):]
                base = letter if len(letter) == 1 else '\\' + letter.capitalize()
            elif rest.startswith('SMALL '):
                letter = rest[len('SMALL '):]
                base = letter.lower() if len(letter) == 1 else '\\' + letter.lower()
            elif rest.startswith('DIGIT '):
                base = str(unicodedata.digit(ch))
            else:
                return ch
            return template % base
    return ch


def _fix_in_text(text: str, custom: list) -> str:
    # 𝑛 (MATHEMATICAL ITALIC SMALL N) in a sentence -> $n$
    text = _MATH_ALNUM_RUN.sub(
        lambda m: '$' + ''.join(_math_alnum_latex(c) for c in m.group(0)) + '$', text)

    # Handle ℓ followed by subscript digits as a unit: ℓ₁ → $\ell_1$
    def _ell_sub(m: re.Match) -> str:
        digits = m.group(1).translate(_SUB_DIGIT_MAP)
        return f'$\\ell_{{{digits}}}$'

    text = _ELL_SUB_TEXT.sub(_ell_sub, text)

    # Lone ℓ → $\ell$
    text = _ELL_TEXT.sub(r'$\\ell$', text)

    # Remaining subscript digits → $_{N}$
    def _sub_digit(m: re.Match) -> str:
        digits = m.group(1).translate(_SUB_DIGIT_MAP)
        return f'$_{{{digits}}}$'

    text = _SUB_DIGIT_TEXT.sub(_sub_digit, text)

    for old, new in _IN_TEXT:
        text = text.replace(old, new)
    for rule in custom:
        if 'char' in rule and 'text' in rule:
            text = text.replace(rule['char'], rule['text'])
    return text
