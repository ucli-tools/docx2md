"""
Front matter structure detection for docx2md.

Detects body content that represents front matter pages (dedication, copyright,
title page repeats) in the pre-heading area of a converted .docx, and wraps
them with special markdown headings (``# Dedication``, ``# Copyright Page``)
so that mdtexpdf's Lua filter can render them as proper front matter pages.

Content between the YAML close (``---``) and the first real ``# `` heading is
split on page break markers (Word section breaks converted by pandoc as
``*\\`` + ``*`` on adjacent lines).  Each section is classified by heuristics:

- **Copyright**: contains "Copyright", "ISBN", "All rights reserved", etc.
- **Dedication**: all-italic short paragraphs, no structural keywords
- **Title page repeat**: contains the book title + author → stripped entirely
"""

import re
import unicodedata
from typing import Any, Dict, List, Optional, Tuple

from docx2md.processors.base import BaseProcessor
from docx2md.utils.logging_utils import get_logger

logger = get_logger(__name__)

# Page break markers produced by pandoc from Word section breaks:
#   *\
#   *
# (with possible trailing whitespace)
_PAGE_BREAK = re.compile(r'\*\\\s*\n\*\s*\n')

# Keywords that indicate a copyright / publisher information page
_COPYRIGHT_KEYWORDS = re.compile(
    r'Copyright|ISBN|All rights reserved|Published by',
    re.IGNORECASE,
)

# An all-italic line: *text* (not bold **text**)
_ITALIC_LINE = re.compile(r'^\*([^*\n]+)\*\s*$')

# Image reference (publisher logos, etc.)
_IMAGE_REF = re.compile(r'!\[.*?\]\(.*?\)')

# YAML front matter block at the start of the document
_YAML_BLOCK = re.compile(r'\A---\n.*?\n---\n', re.DOTALL)

# First real heading (# at level 1 or 2)
_FIRST_HEADING = re.compile(r'^#{1,2}\s+', re.MULTILINE)

# Standalone "Volume X, Part Y:" line (Word title page artifact)
_VOLUME_PART_LINE = re.compile(
    r'^\s*\*{0,3}Volume\s+[IVXLCDM]+,?\s*Part\s+\d+\s*:\s*\*{0,3}\s*$',
    re.IGNORECASE,
)


class FrontMatterStructureProcessor(BaseProcessor):
    """Detect and structure body front matter from converted .docx files."""

    def process(self, content: str, doc_properties: Optional[Dict[str, Any]] = None) -> str:
        pre, heading_start = self._split_pre_heading(content)
        if not pre.strip():
            return content

        # Extract title and author for title-page-repeat detection
        # Priority: doc_properties (from python-docx) > YAML > body scan
        title, author = self._extract_metadata(content, doc_properties or {})

        sections = _PAGE_BREAK.split(pre)

        classified: List[Tuple[str, str]] = []  # (type, text)
        for section in sections:
            text = section.strip()
            if not text:
                continue
            kind = self._classify(text, title, author)
            classified.append((kind, text))

        if not classified:
            return content

        # If everything classified as 'unknown', nothing to restructure
        if all(kind == 'unknown' for kind, _ in classified):
            return content

        # Build replacement for the pre-heading area
        parts: List[str] = []
        for kind, text in classified:
            if kind == 'title_repeat':
                continue  # strip entirely
            elif kind == 'copyright':
                parts.append(f'# Copyright Page\n\n{text}')
            elif kind == 'dedication':
                parts.append(f'# Dedication\n\n{text}')
            else:
                # Unknown — keep as-is
                parts.append(text)

        new_pre = '\n\n'.join(parts)
        if new_pre:
            new_pre += '\n\n'

        # Reconstruct: everything before pre + new_pre + heading onward
        result = self._reconstruct(content, new_pre, heading_start)

        # Second pass: strip orphaned title/subtitle fragments in the body
        # (e.g. "Volume I, Part 1:\n\nMathematical Foundations\n\nand the Singularity")
        if title:
            result = self._strip_body_title_fragments(result, title)

        return result

    def _split_pre_heading(self, content: str) -> Tuple[str, int]:
        """Return (pre-heading text, index of first heading) from body content.

        The "body" starts after the YAML front matter block (if present).
        """
        body_start = 0
        yaml_match = _YAML_BLOCK.match(content)
        if yaml_match:
            body_start = yaml_match.end()

        body = content[body_start:]
        heading_match = _FIRST_HEADING.search(body)
        if heading_match:
            pre = body[:heading_match.start()]
            heading_start = body_start + heading_match.start()
        else:
            # No heading found — entire body is pre-heading
            pre = body
            heading_start = len(content)

        return pre, heading_start

    def _reconstruct(self, content: str, new_pre: str, heading_start: int) -> str:
        """Rebuild the document with new pre-heading content."""
        yaml_match = _YAML_BLOCK.match(content)
        yaml_part = content[:yaml_match.end()] if yaml_match else ''
        heading_part = content[heading_start:]
        return yaml_part + new_pre + heading_part

    def _extract_metadata(
        self, content: str, doc_properties: Dict[str, Any],
    ) -> Tuple[str, str]:
        """Extract title and author from available sources.

        Priority order:
        1. ``doc_properties`` from python-docx (Step 1 metadata)
        2. YAML front matter (if present)
        3. Body bold-title scan (fallback)
        """
        title = doc_properties.get('title', '').strip()
        author = doc_properties.get('author', '').strip()

        # Try YAML if doc_properties didn't have them
        if not title or not author:
            yaml_match = _YAML_BLOCK.match(content)
            if yaml_match:
                yaml_text = yaml_match.group()
                if not title:
                    m = re.search(r'^title:\s*"?([^"\n]+)"?\s*$', yaml_text, re.MULTILINE)
                    if m:
                        title = m.group(1).strip()
                if not author:
                    m = re.search(r'^author:\s*"?([^"\n]+)"?\s*$', yaml_text, re.MULTILINE)
                    if m:
                        author = m.group(1).strip()

        # Fallback: scan body for bold title
        if not title:
            m = re.search(r'^\*\*([^*\n]+)\*\*\s*$', content[:500], re.MULTILINE)
            if m:
                title = m.group(1).strip()

        return title, author

    def _classify(self, text: str, title: str, author: str) -> str:
        """Classify a front matter section.

        Returns one of: 'copyright', 'dedication', 'title_repeat', 'unknown'.
        """
        # Copyright page: contains copyright keywords
        if _COPYRIGHT_KEYWORDS.search(text):
            return 'copyright'

        # Title page repeat: contains both the book title and author name
        # Uses accent-insensitive comparison (NFD decomposition strips accents)
        # Also tries "First Last" when author is "Last, First"
        if title and author:
            text_norm = self._strip_accents(text.lower())
            has_title = self._strip_accents(title.lower()) in text_norm
            author_norm = self._strip_accents(author.lower())
            has_author = author_norm in text_norm
            # Try flipped name: "Doe, Jane" → "Jane Doe"
            if not has_author and ',' in author:
                parts = author.split(',', 1)
                if len(parts) == 2 and parts[1].strip():
                    flipped = f"{parts[1].strip()} {parts[0].strip()}"
                    has_author = self._strip_accents(flipped.lower()) in text_norm
            if has_title and has_author:
                return 'title_repeat'

        # Title fragment: contains the title, short, and all lines are
        # formatted (italic/bold) — a subtitle or volume page without author
        if title:
            text_norm = self._strip_accents(text.lower())
            lines = [line for line in text.split('\n') if line.strip()]
            if (self._strip_accents(title.lower()) in text_norm
                    and len(lines) <= 5
                    and all(self._is_formatted_or_empty(line) for line in lines)):
                return 'title_repeat'

        # Dedication: all non-empty lines are italic, no images, short
        lines = [line for line in text.split('\n') if line.strip()]
        if lines and all(self._is_italic_or_empty(line) for line in lines):
            # Exclude if it has images (publisher logo page)
            if not _IMAGE_REF.search(text):
                return 'dedication'

        return 'unknown'

    def _strip_body_title_fragments(self, content: str, title: str) -> str:
        """Strip orphaned title/subtitle/volume page fragments from the body.

        Word title pages sometimes produce fragments like::

            Volume I, Part 1:

            Mathematical Foundations

            and the Singularity

        These appear between headings as standalone short paragraphs.
        Anchored on "Volume X, Part Y:" lines, removes the anchor and
        adjacent short subtitle paragraphs.
        """
        lines = content.split('\n')

        # Find "Volume X, Part Y:" anchor lines
        anchors = [i for i, line in enumerate(lines) if _VOLUME_PART_LINE.match(line)]
        if not anchors:
            return content

        to_remove = set()
        for anchor in anchors:
            to_remove.add(anchor)

            # Remove short subtitle lines forward until heading or long paragraph
            i = anchor + 1
            while i < len(lines):
                stripped = lines[i].strip()
                if stripped.startswith('#'):
                    break
                if not stripped:
                    to_remove.add(i)
                    i += 1
                    continue
                # Short unformatted line (< 80 chars, no images, no math)
                if len(stripped) < 80 and not stripped.startswith('!') and '$' not in stripped:
                    to_remove.add(i)
                    i += 1
                else:
                    break

            # Remove blank lines backward
            i = anchor - 1
            while i >= 0 and not lines[i].strip():
                to_remove.add(i)
                i -= 1

        result = [line for i, line in enumerate(lines) if i not in to_remove]
        # Ensure blank line before any heading that now directly follows text
        final = []
        for j, line in enumerate(result):
            if (line.startswith('#') and j > 0
                    and result[j - 1].strip()):
                final.append('')
            final.append(line)
        return '\n'.join(final)

    @staticmethod
    def _strip_accents(s: str) -> str:
        """Remove accent marks for fuzzy title matching."""
        return ''.join(
            c for c in unicodedata.normalize('NFD', s)
            if unicodedata.category(c) != 'Mn'
        )

    @staticmethod
    def _is_italic_or_empty(line: str) -> bool:
        """Check if a line is all-italic markdown or empty."""
        stripped = line.strip()
        if not stripped:
            return True
        return bool(_ITALIC_LINE.match(stripped))

    @staticmethod
    def _is_formatted_or_empty(line: str) -> bool:
        """Check if a line is italic, bold, bold-italic (possibly nested), or empty."""
        stripped = line.strip()
        if not stripped:
            return True
        # Line starts with * and ends with * (covers *italic*, **bold**,
        # ***bold-italic***, and nested like *text **inner** text*)
        return stripped.startswith('*') and stripped.rstrip().endswith('*')
