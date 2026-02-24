"""
YAML frontmatter generator for docx2md.

Generates mdtexpdf-compatible YAML frontmatter from document metadata
and prepends it to the markdown content.

The generated frontmatter includes:
  - Active fields: auto-detected from the .docx (title, author, subtitle)
  - Active defaults: format, toc, pagination settings
  - Commented template: all mdtexpdf fields organized by section, ready
    for editorial customization (dedication, covers, publisher, etc.)

Metadata sources (priority order):
  1. python-docx core properties (title, author, subject)
  2. Body text scan: **Bold Title**, *Italic subtitle*, Name -- <email>
  3. Config overrides (config["yaml_frontmatter"]["mdtexpdf"])
  4. Input filename stem as last-resort title
"""

import re
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

from docx2md.utils.logging_utils import get_logger

logger = get_logger(__name__)

# Matches a bold-only first paragraph: **Title text**
_BOLD_TITLE = re.compile(r'^\*\*([^*\n]+)\*\*\s*$', re.MULTILINE)

# Matches an italic-only paragraph: *Subtitle text*
_ITALIC_SUBTITLE = re.compile(r'^\*([^*\n]+)\*\s*$', re.MULTILINE)

# Matches "Name -- <email>" or "Name -- email@..." author line
_AUTHOR_LINE = re.compile(
    r'^([A-Z][^\n<]*?)\s+--\s+<?([a-zA-Z0-9._%+\-]+@[^\s>]+)>?\s*$',
    re.MULTILINE,
)

# Title block at start of document: optional bold title + optional italic subtitle +
# optional author line, each separated by a blank line
_TITLE_BLOCK = re.compile(
    r'^\*\*[^*\n]+\*\*\s*\n\n'          # bold title + blank line
    r'(?:\*[^*\n]+\*\s*\n\n)?'           # optional italic subtitle + blank line
    r'(?:[A-Z][^\n<]*?--[^\n]+\n\n)?',   # optional author line + blank line
    re.MULTILINE,
)

# Headings that should get \newpage before them
_NEWPAGE_HEADINGS = re.compile(
    r'^(#{1,2}\s+(?:References|Bibliography|Index|Glossary)\s*)$',
    re.MULTILINE,
)


def _escape_yaml(s: str) -> str:
    """Escape a string for use inside YAML double quotes."""
    return s.replace('\\', '\\\\').replace('"', '\\"')


def _fmt(value: Any) -> str:
    """Format a Python value for YAML output."""
    if isinstance(value, bool):
        return 'true' if value else 'false'
    if isinstance(value, int):
        return str(value)
    if isinstance(value, float):
        return str(value)
    if isinstance(value, str):
        return f'"{_escape_yaml(value)}"'
    return str(value)


def _flip_author_name(name: str) -> str:
    """Flip 'Last, First' or 'Last, First Middle' to 'First Middle Last'."""
    if ',' in name:
        parts = name.split(',', 1)
        if len(parts) == 2 and parts[1].strip():
            return f"{parts[1].strip()} {parts[0].strip()}"
    return name


def generate_yaml_frontmatter(
    doc_properties: Dict[str, Any],
    config: Dict[str, Any],
    input_path: Path,
    content: str,
) -> Tuple[str, str]:
    """
    Generate a comprehensive YAML frontmatter block for mdtexpdf.

    Auto-detected fields are active; book-specific fields are commented
    out with descriptions so an editor can enable them.

    Args:
        doc_properties: Metadata extracted by python-docx (title, author, subject, ...)
        config:         Full docx2md config dict
        input_path:     Path to the original .docx file (used for title fallback)
        content:        Current markdown content (may be scanned for title/author)

    Returns:
        Tuple of (frontmatter_string, updated_content) where frontmatter_string is
        the YAML block (including --- delimiters) and updated_content has the
        title block removed and \\newpage inserted before References/Index.
    """
    fm_cfg = config.get('yaml_frontmatter', {})

    if not fm_cfg.get('enabled', True):
        return '', content

    # Guard: don't double-prepend if content already starts with YAML
    if content.lstrip().startswith('---'):
        logger.debug('Content already has YAML frontmatter, skipping generation')
        return '', content

    # ---- Collect metadata ----
    title = doc_properties.get('title', '').strip()
    author = doc_properties.get('author', '').strip()
    subtitle = doc_properties.get('subject', '').strip()
    email = ''

    # Scan body text if allowed and properties are incomplete
    if fm_cfg.get('extract_from_body', True):
        title, subtitle, author, email = _extract_from_body(
            content, title, subtitle, author
        )

    # Last resort: derive title from filename
    if not title:
        title = input_path.stem.replace('_', ' ').replace('-', ' ').title()

    # Fall back to config default_author
    if not author:
        author = fm_cfg.get('default_author', '')

    # Flip "Last, First" → "First Last"
    author = _flip_author_name(author)

    # ---- Config overrides ----
    overrides = dict(fm_cfg.get('mdtexpdf', {}))

    # ---- Detect body front matter headings ----
    has_body_frontmatter = {
        'dedication': bool(re.search(r'^# Dedication\s*$', content, re.MULTILINE)),
        'copyright_page': bool(re.search(r'^# Copyright Page\s*$', content, re.MULTILINE)),
    }

    # ---- Build template ----
    now = datetime.now()
    date_str = now.strftime('%B %Y')
    copyright_year = now.year

    frontmatter = _build_yaml_template(
        title=title,
        subtitle=subtitle,
        author=author,
        email=email,
        date_str=date_str,
        copyright_year=copyright_year,
        overrides=overrides,
        has_body_frontmatter=has_body_frontmatter,
    )

    # ---- Strip title block from body ----
    updated_content = content
    if fm_cfg.get('strip_body_title_block', True):
        updated_content = _strip_title_block(content)

    # ---- Insert \newpage before References/Index headings ----
    updated_content = _insert_newpage_before_sections(updated_content)

    logger.debug(f'Generated frontmatter for: {title!r} by {author!r}')
    return frontmatter, updated_content


def _build_yaml_template(
    title: str,
    subtitle: str,
    author: str,
    email: str,
    date_str: str,
    copyright_year: int,
    overrides: Optional[Dict[str, Any]] = None,
    has_body_frontmatter: Optional[Dict[str, bool]] = None,
) -> str:
    """
    Build comprehensive YAML frontmatter with commented editorial fields.

    Active fields come from auto-detection + overrides.
    Commented fields provide a ready-to-edit template for the editor.

    When ``has_body_frontmatter`` indicates that a ``# Dedication`` or
    ``# Copyright Page`` heading exists in the markdown body, the
    corresponding YAML field is always commented out (even if it would
    otherwise be force_active) to prevent duplication.
    """
    if overrides is None:
        overrides = {}
    if has_body_frontmatter is None:
        has_body_frontmatter = {}

    lines: List[str] = []
    lines.append('---')

    # Helper: emit an active (uncommented) key-value pair
    def active(key: str, value: Any) -> None:
        lines.append(f'{key}: {_fmt(value)}')

    # Helper: emit a commented key-value pair
    def commented(key: str, value: Any) -> None:
        lines.append(f'# {key}: {_fmt(value)}')

    # Helper: emit a field — active if in overrides or if force_active, else commented
    def field(key: str, default: Any, force_active: bool = False) -> None:
        if key in overrides:
            active(key, overrides[key])
        elif force_active:
            active(key, default)
        else:
            commented(key, default)

    # === COMMON METADATA ===
    lines.append('# === COMMON METADATA ===')
    active('title', title)
    if subtitle or 'subtitle' in overrides:
        active('subtitle', overrides.get('subtitle', subtitle))
    else:
        commented('subtitle', 'Subtitle goes here')
    active('author', overrides.get('author', author))
    if email or 'email' in overrides:
        active('email', overrides.get('email', email))
    else:
        commented('email', 'author@example.com')
    active('date', overrides.get('date', date_str))
    field('description', 'Brief description of this book.')
    lines.append('')

    # === PDF SETTINGS ===
    lines.append('# === PDF SETTINGS ===')
    field('format', 'book', force_active=True)
    field('no_numbers', True, force_active=True)
    field('toc', True, force_active=True)
    field('lof', True)
    field('lot', True)
    field('index', True)
    field('header_footer_policy', 'all', force_active=True)
    footer_default = f'\u00a9 {copyright_year} {author}. All rights reserved.'
    field('footer', footer_default)
    field('pageof', True, force_active=True)
    field('date_footer', True, force_active=True)
    lines.append('')

    # === PROFESSIONAL BOOK FEATURES ===
    lines.append('# === PROFESSIONAL BOOK FEATURES ===')
    field('half_title', False)
    # If body already has # Copyright Page / # Dedication headings, comment out
    # the YAML fields to prevent duplication
    if has_body_frontmatter.get('copyright_page'):
        commented('copyright_page', True)
        lines.append('# (detected # Copyright Page heading in body)')
    else:
        field('copyright_page', True, force_active=True)
    if has_body_frontmatter.get('dedication'):
        commented('dedication', 'To whom this book is dedicated.')
        lines.append('# (detected # Dedication heading in body)')
    else:
        field('dedication', 'To whom this book is dedicated.')
    field('epigraph', 'An inspiring quote.')
    field('epigraph_source', 'Author of the quote')
    field('chapters_on_recto', True, force_active=True)
    field('drop_caps', True, force_active=True)
    field('equation_numbers', True)
    field('publisher', 'Publisher Name')
    field('copyright_year', copyright_year, force_active=True)
    field('edition', 'First Edition', force_active=True)
    field('edition_date', date_str)
    field('printing', f'First Printing, {date_str}')
    field('publisher_address', 'Address')
    field('publisher_website', 'www.publisher.com')
    lines.append('')

    # === COVER SYSTEM ===
    lines.append('# === COVER SYSTEM ===')
    field('cover_image', 'img/cover.jpeg')
    field('cover_title_color', 'white')
    field('cover_title_show', True)
    field('cover_subtitle_show', True)
    field('cover_author_position', 'bottom')
    field('cover_overlay_opacity', 0.4)
    field('cover_fit', 'cover')
    lines.append('#')
    field('back_cover_image', 'img/back.jpeg')
    field('back_cover_content', 'quote')
    field('back_cover_text', 'Back cover description text.')
    field('back_cover_author_bio', True)
    field('back_cover_author_bio_text', 'Author bio text.')
    field('back_cover_isbn_barcode', False)
    field('back_cover_text_background', True)
    field('back_cover_text_background_opacity', 0.3)
    field('back_cover_text_color', 'white')
    lines.append('')

    # === PRINT FORMAT ===
    lines.append('# === PRINT FORMAT ===')
    field('trim_size', '6x9')
    field('paper_stock', 'cream60')
    field('spine_text', 'auto')
    lines.append('')

    # === ACKNOWLEDGMENTS & ABOUT THE AUTHOR ===
    lines.append('# === ACKNOWLEDGMENTS & ABOUT THE AUTHOR ===')
    if 'acknowledgments' in overrides:
        lines.append('acknowledgments: |')
        for ack_line in str(overrides['acknowledgments']).split('\n'):
            lines.append(f'  {ack_line}')
    else:
        lines.append('# acknowledgments: |')
        lines.append('#   Acknowledgment text here.')
    if 'about-author' in overrides:
        active('about-author', overrides['about-author'])
    else:
        commented('about-author', 'Author bio for the About the Author page.')
    lines.append('')

    # === LATEX PACKAGES ===
    lines.append('# === LATEX PACKAGES (uncomment if book uses TikZ/pgfplots diagrams) ===')
    if 'header-includes' in overrides:
        lines.append('header-includes:')
        for inc in overrides['header-includes']:
            lines.append(f'  - {inc}')
    else:
        lines.append('# header-includes:')
        lines.append('#   - \\usepackage{pgfplots}')
        lines.append('#   - \\pgfplotsset{compat=1.17}')
        lines.append('#   - \\usetikzlibrary{arrows.meta,decorations.markings,calc}')
        lines.append('#   - \\usepgfplotslibrary{fillbetween}')
    lines.append('')

    lines.append('---')
    lines.append('')

    return '\n'.join(lines) + '\n'


def _insert_newpage_before_sections(content: str) -> str:
    """Insert \\newpage before References, Bibliography, Index, Glossary headings."""
    def _add_newpage(m: re.Match) -> str:
        return f'\\newpage\n\n{m.group(1)}'
    return _NEWPAGE_HEADINGS.sub(_add_newpage, content)


def _extract_from_body(
    content: str,
    title: str,
    subtitle: str,
    author: str,
) -> Tuple[str, str, str, str]:
    """
    Scan the first paragraphs of the body for title/subtitle/author patterns.
    Only fills in values that are still empty.
    """
    email = ''

    # Look for bold title paragraph
    if not title:
        m = _BOLD_TITLE.search(content[:500])
        if m:
            title = m.group(1).strip()

    # Look for italic subtitle paragraph
    if not subtitle:
        m = _ITALIC_SUBTITLE.search(content[:800])
        if m:
            subtitle = m.group(1).strip()

    # Look for "Name -- <email>" author line
    if not author:
        m = _AUTHOR_LINE.search(content[:1000])
        if m:
            author = m.group(1).strip()
            email = m.group(2).strip()

    return title, subtitle, author, email


def _strip_title_block(content: str) -> str:
    """
    Remove the leading title/subtitle/author paragraphs from the body
    (they have been moved into the YAML frontmatter).
    """
    return _TITLE_BLOCK.sub('', content, count=1)
