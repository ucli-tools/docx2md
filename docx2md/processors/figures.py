"""
Figure caption processor for docx2md.

Fixes the double-caption problem caused by Word documents having both:
  - An AI-generated alt-text on the image (captured by pandoc as the [alt] label)
  - A real figure caption paragraph immediately following the image

This processor:
  1. Detects image blocks followed by a "Figure N." caption paragraph
  2. Replaces the image's generic alt-text with the real caption
  3. Removes the now-redundant caption paragraph
"""

import re
from typing import Any, Dict, List, Optional, Tuple

from docx2md.processors.base import BaseProcessor
from docx2md.utils.logging_utils import get_logger

logger = get_logger(__name__)

# Matches the start of a figure caption paragraph.
# Handles various Word rendering artifacts:
#   ***Figure 1.***  →  tight bold-italic
#   *Figure* *1.*   →  split italic (Word sometimes splits number/label)
#   ***Figure* *4.*-**  →  mixed asterisks
# Strategy: look for "Figure" anywhere near the start (within first 30 chars)
_CAPTION_START = re.compile(
    r'^\*{1,3}Figure',
    re.IGNORECASE,
)

# Matches an image reference: ![...](path)  (possibly multi-line alt text)
_IMAGE_START = re.compile(r'^!\[')
_IMAGE_ANY = re.compile(r'!\[')

# Strips emphasis markers and collapses whitespace from caption text
_EMPHASIS = re.compile(r'\*{1,3}(.*?)\*{1,3}', re.DOTALL)
_NEWLINE_SPACE = re.compile(r'\s*\n\s*')


class FigureProcessor(BaseProcessor):
    """
    Replaces AI-generated image alt-text with real figure captions.

    Args:
        config: Configuration dict (uses config["figures"] section)
    """

    def __init__(self, config: Optional[Dict[str, Any]] = None):
        super().__init__(config)
        cfg = (config or {}).get('figures', {})
        self.enabled = cfg.get('enabled', True)

    def process(self, content: str) -> str:
        if not self.enabled:
            return content

        # Split on blank lines to get paragraph blocks
        blocks = re.split(r'\n{2,}', content)
        result: List[str] = []
        i = 0

        while i < len(blocks):
            block = blocks[i].strip()

            if _IMAGE_START.match(block):
                # Look ahead for the next non-empty block
                j = i + 1
                while j < len(blocks) and not blocks[j].strip():
                    j += 1

                if j < len(blocks):
                    next_block = blocks[j].strip()
                    cap_match = _CAPTION_START.match(next_block)

                    if cap_match:
                        # Extract and clean the real caption text
                        caption_text = _extract_caption(next_block)
                        if len(_IMAGE_ANY.findall(block)) > 1:
                            # Images side by side share one caption. pandoc
                            # captions only an image alone in its paragraph,
                            # so keep the images together and the caption as
                            # its own paragraph beneath them
                            result.append(_replace_alt_text(block, ''))
                            result.append(f'*{caption_text}*')
                        else:
                            # Replace AI alt-text in the image block with the real caption
                            result.append(_replace_alt_text(block, caption_text))
                        # Skip over any blank blocks and the caption paragraph
                        i = j + 1
                        logger.debug(f'Replaced figure caption: {caption_text[:60]}...')
                        continue

            result.append(blocks[i])
            i += 1

        return '\n\n'.join(result)


def _extract_caption(caption_block: str) -> str:
    """
    Extract clean plain text (+math) from a figure caption paragraph.

    Input example:
        ***Figure 1.*** *-- Polyplex Lie Geometry image* $\\in \\mathcal{H}^3$*...*

    Output:
        Figure 1. -- Polyplex Lie Geometry image $\\in \\mathcal{H}^3$...
    """
    text = caption_block

    # Collapse line-wrap newlines inside the paragraph
    text = _NEWLINE_SPACE.sub(' ', text)

    # Strip outermost ***...*** or **...** or *...* wrapping iteratively
    # but preserve math spans untouched
    # Strategy: strip emphasis markers that are NOT inside $...$
    text = _strip_emphasis_outside_math(text)

    # Sanitize math for use inside \caption{} (LaTeX-safe)
    text = _sanitize_caption_math(text)

    return text.strip()


def _sanitize_caption_math(text: str) -> str:
    """
    Make math spans safe for LaTeX \\caption{} context.

    Handles both pandoc delimiter styles:
    - Display math: $$ or \\[...\\] → convert to inline $
    - Inline math: $ or \\(...\\) → keep as $
    - Strip any math span with unbalanced braces (garbled Word output)
    """
    # Convert \[...\] display math to $...$ inline (display illegal in captions)
    text = re.sub(r'\\\[', '$', text)
    text = re.sub(r'\\\]', '$', text)
    # Convert \(...\) to $...$ for uniformity
    text = re.sub(r'\\\(', '$', text)
    text = re.sub(r'\\\)', '$', text)
    # Convert remaining $$ → $ (display math cannot appear inside \caption{})
    text = text.replace('$$', '$')

    # Walk through math spans and drop any with unbalanced braces
    result = []
    i = 0
    while i < len(text):
        if text[i] == '$' and (i == 0 or text[i - 1] != '\\'):
            j = text.find('$', i + 1)
            if j == -1:
                break  # no closing $ — drop remainder
            span = text[i + 1:j]
            if span.count('{') == span.count('}'):
                result.append(f'${span}$')
            # else: skip broken math span
            i = j + 1
        else:
            result.append(text[i])
            i += 1

    return ''.join(result)


def _strip_emphasis_outside_math(text: str) -> str:
    """
    Remove *...* and ***...*** markers from text, leaving $math$ spans intact.
    """
    # Tokenize math vs text
    from docx2md.utils.math_utils import tokenize_math_spans
    tokens = tokenize_math_spans(text)
    result = []
    for kind, part in tokens:
        if kind == 'text':
            # Remove emphasis markers
            part = re.sub(r'(?<!\\)\*{1,3}(.*?)(?<!\\)\*{1,3}', r'\1', part, flags=re.DOTALL)
            # An italic run that spans a math span leaves one marker on each
            # side of it, unpaired within its text piece: drop what is left
            # (a literal asterisk arrives escaped as \* and is kept)
            part = re.sub(r'(?<!\\)\*+', '', part)
            # Clean up double spaces
            part = re.sub(r'  +', ' ', part)
        result.append(part)
    return ''.join(result)


def _replace_alt_text(image_block: str, new_alt: str) -> str:
    """
    Replace the [alt text] portion in all image references within a block.

    Handles multi-line alt text that pandoc may have word-wrapped.
    """
    # Pattern matches ![...multi...line...](path)
    # We replace the [...] part (everything between ![ and ](
    replacement = f'![{new_alt}]'
    result = re.sub(
        r'!\[[\s\S]*?\](?=\()',  # match ![ ... ] (non-greedy, up to the ()
        lambda _: replacement,   # lambda avoids re treating \cmd as backrefs
        image_block,
    )
    return result
