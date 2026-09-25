"""
Placeholder-based math extraction for docx2md.

Replaces the single-pandoc-call approach with a 3-phase pipeline:

1. **Extract** — Parse .docx XML, pull out all ``<m:oMathPara>`` (display)
   and ``<m:oMath>`` (inline) elements, replace each with a unique text
   placeholder, and re-zip into a sanitised .docx.
2. **Convert** — Run pandoc on the math-free .docx (structure only) **and**
   batch-convert the extracted equations via a separate pandoc call.
3. **Splice** — Replace placeholders in the markdown with properly
   delimited LaTeX (``$...$`` for inline, ``$$...$$`` for display).

This eliminates pandoc's intermittent dropping of ``$`` delimiters around
OMML-converted equations.
"""

import re
import shutil
import tempfile
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple, Union
from xml.etree import ElementTree as ET

import pypandoc

from docx2md.utils.docx_xml_utils import (
    NAMESPACES,
    create_text_run,
    register_omml_namespaces,
    rezip_docx,
    unzip_docx,
)
from docx2md.utils.logging_utils import get_logger

logger = get_logger(__name__)

# Namespace URIs wrapped in braces for ElementTree tag construction
_NS_W = "{" + NAMESPACES["w"] + "}"
_NS_M = "{" + NAMESPACES["m"] + "}"


class MathExtractor:
    """Extract math from .docx XML, convert via pandoc, splice back."""

    def __init__(self, config: Optional[Dict[str, Any]] = None):
        self.config = config or {}
        # Word index entries (XE fields) found in phase 1, spliced in phase 3
        self._index_fields: List[Dict[str, Any]] = []

    # ------------------------------------------------------------------
    # Public API
    # ------------------------------------------------------------------

    def extract_and_convert(
        self,
        docx_path: Union[str, Path],
        media_dir: Union[str, Path],
        extra_args: List[str],
    ) -> Tuple[str, Dict[str, Any]]:
        """Top-level orchestrator.

        Args:
            docx_path: Path to the original .docx file.
            media_dir: Directory for extracted media (images).
            extra_args: Extra arguments to pass to pandoc.

        Returns:
            (markdown_text, stats_dict)
        """
        docx_path = Path(docx_path)
        media_dir = Path(media_dir)

        with tempfile.TemporaryDirectory(prefix="docx2md_math_") as tmp_dir:
            tmp = Path(tmp_dir)

            # Phase 1: extract math, create sanitised docx
            sanitized_docx, equations = self._extract_math_from_docx(
                docx_path, tmp
            )

            logger.info(
                "Extracted %d equations (%d display, %d inline)",
                len(equations),
                sum(1 for e in equations if e["kind"] == "display"),
                sum(1 for e in equations if e["kind"] == "inline"),
            )

            # Phase 2a: pandoc on math-free docx (structure + text)
            markdown = self._run_pandoc(sanitized_docx, media_dir, extra_args)

            # Phase 2b: batch-convert equations to LaTeX
            if equations:
                eq_latex = self._batch_convert_equations(equations, tmp)
            else:
                eq_latex = {}

            # Phase 3: splice equations back into markdown
            markdown = self._splice(markdown, eq_latex, equations)
            markdown = self._splice_index(markdown, eq_latex)

        stats = {
            "index_entries": len(self._index_fields),
            "math_equations_extracted": len(equations),
            "math_display_count": sum(
                1 for e in equations if e["kind"] == "display"
            ),
            "math_inline_count": sum(
                1 for e in equations if e["kind"] == "inline"
            ),
        }
        return markdown, stats

    # ------------------------------------------------------------------
    # Phase 1: XML extraction
    # ------------------------------------------------------------------

    def _extract_math_from_docx(
        self, docx_path: Path, tmp_dir: Path
    ) -> Tuple[Path, List[Dict[str, Any]]]:
        """Parse document.xml, replace math elements with placeholders.

        Returns:
            (path_to_sanitized_docx, list_of_equation_dicts)

        Each equation dict has keys: ``idx``, ``kind`` (inline|display),
        ``xml`` (serialised element bytes), ``placeholder``.
        """
        register_omml_namespaces()

        unpack_dir = tmp_dir / "unpack"
        unzip_docx(docx_path, unpack_dir)

        doc_xml_path = unpack_dir / "word" / "document.xml"
        tree = ET.parse(doc_xml_path)
        root = tree.getroot()

        # Build child→parent map (ElementTree has no parent pointers)
        parent_map: Dict[ET.Element, ET.Element] = {}
        for parent in root.iter():
            for child in parent:
                parent_map[child] = parent

        # Word index entries become [index:...] markers (their math is
        # converted with the other equations, below)
        self._index_fields = []
        if self.config.get("processing", {}).get("index_entries", True):
            self._index_fields = self._collect_index_fields(root, parent_map)

        # Math inside a field instruction (an XE index entry, a TC entry) is
        # part of the hidden field code, not of the text: drop it
        for math_el in self._math_in_field_instructions(root, parent_map):
            parent = parent_map.get(math_el)
            if parent is not None:
                parent.remove(math_el)

        equations: List[Dict[str, Any]] = []
        idx = 0

        # Process display equations FIRST (oMathPara contains inner oMath)
        for math_para in list(root.iter(f"{_NS_M}oMathPara")):
            placeholder = f"@@MATH_DISPLAY_{idx:04d}@@"
            xml_bytes = ET.tostring(math_para, encoding="unicode")

            equations.append({
                "idx": idx,
                "kind": "display",
                "xml": xml_bytes,
                "placeholder": placeholder,
            })

            parent = parent_map.get(math_para)
            if parent is not None:
                pos = list(parent).index(math_para)
                parent.remove(math_para)
                run = create_text_run(placeholder, _NS_W)
                parent.insert(pos, run)
                # Update parent map for new element
                parent_map[run] = parent

            idx += 1

        # Process remaining inline oMath (those NOT inside an oMathPara)
        for math_el in list(root.iter(f"{_NS_M}oMath")):
            # Skip if this oMath is inside an oMathPara (already handled)
            ancestor = parent_map.get(math_el)
            inside_para = False
            while ancestor is not None:
                if ancestor.tag == f"{_NS_M}oMathPara":
                    inside_para = True
                    break
                ancestor = parent_map.get(ancestor)
            if inside_para:
                continue

            placeholder = f"@@MATH_INLINE_{idx:04d}@@"
            xml_bytes = ET.tostring(math_el, encoding="unicode")

            equations.append({
                "idx": idx,
                "kind": "inline",
                "xml": xml_bytes,
                "placeholder": placeholder,
            })

            parent = parent_map.get(math_el)
            if parent is not None:
                pos = list(parent).index(math_el)
                parent.remove(math_el)
                run = create_text_run(placeholder, _NS_W)
                parent.insert(pos, run)
                parent_map[run] = parent

            idx += 1

        # Math inside index terms: converted in the same batch, spliced into
        # the [index:...] markers (never into the text)
        for field in self._index_fields:
            for k, part in enumerate(field["parts"]):
                if isinstance(part, str):
                    continue
                equations.append({
                    "idx": idx,
                    "kind": "inline",
                    "xml": ET.tostring(part, encoding="unicode"),
                    "placeholder": f"@@MATH_INDEX_{idx:04d}@@",
                })
                plain = "".join(t.text or "" for t in part.iter(f"{_NS_M}t"))
                field["parts"][k] = ("math", idx, plain)
                idx += 1

        # Write modified XML back
        tree.write(doc_xml_path, xml_declaration=True, encoding="UTF-8")

        # Re-zip
        sanitized_path = tmp_dir / "sanitized.docx"
        rezip_docx(unpack_dir, sanitized_path)

        return sanitized_path, equations

    @staticmethod
    def _math_in_field_instructions(
        root: ET.Element, parent_map: Dict[ET.Element, ET.Element]
    ) -> List[ET.Element]:
        """Math elements between a field's begin and separate (or end) marks.

        Walks the document in order, tracking nested complex fields; returns
        the outermost math elements met while inside an instruction.
        """
        found: List[ET.Element] = []
        stack: List[str] = []
        for el in root.iter():
            if el.tag == f"{_NS_W}fldChar":
                kind = el.get(f"{_NS_W}fldCharType")
                if kind == "begin":
                    stack.append("instruction")
                elif kind == "separate" and stack:
                    stack[-1] = "result"
                elif kind == "end" and stack:
                    stack.pop()
            elif el.tag in (f"{_NS_M}oMath", f"{_NS_M}oMathPara"):
                if stack and stack[-1] == "instruction":
                    found.append(el)
        chosen = set(found)

        def inside_chosen(el: ET.Element) -> bool:
            parent = parent_map.get(el)
            while parent is not None:
                if parent in chosen:
                    return True
                parent = parent_map.get(parent)
            return False

        return [el for el in found if not inside_chosen(el)]

    # ------------------------------------------------------------------
    # Word index entries (XE fields)
    # ------------------------------------------------------------------

    def _collect_index_fields(
        self, root: ET.Element, parent_map: Dict[ET.Element, ET.Element]
    ) -> List[Dict[str, Any]]:
        """Find XE fields and leave a text placeholder where each one ends.

        Each returned field has ``placeholder`` and ``parts``: the pieces of
        its instruction in order, strings from ``w:instrText`` and ``m:oMath``
        elements for math typed into the term.
        """
        fields: List[Dict[str, Any]] = []
        stack: List[Dict[str, Any]] = []
        for el in list(root.iter()):
            if el.tag == f"{_NS_W}fldChar":
                kind = el.get(f"{_NS_W}fldCharType")
                if kind == "begin":
                    stack.append({"state": "instruction", "parts": []})
                elif kind == "separate" and stack:
                    stack[-1]["state"] = "result"
                elif kind == "end" and stack:
                    field = stack.pop()
                    code = "".join(p for p in field["parts"] if isinstance(p, str))
                    if code.lstrip().startswith("XE"):
                        field["end"] = el
                        fields.append(field)
            elif stack and stack[-1]["state"] == "instruction":
                if el.tag == f"{_NS_W}instrText":
                    stack[-1]["parts"].append(el.text or "")
                elif (el.tag == f"{_NS_M}oMath"
                      and parent_map.get(el) is not None
                      and parent_map[el].tag != f"{_NS_M}oMathPara"):
                    stack[-1]["parts"].append(el)

        for n, field in enumerate(fields):
            field["placeholder"] = f"@@INDEX_{n:04d}@@"
            anchor = parent_map.get(field.pop("end"))
            # A field typed inside an equation ends inside it; pandoc's math
            # reader skips text runs there, so anchor after the equation
            node = anchor
            while node is not None:
                if node.tag.startswith(_NS_M):
                    anchor = node
                node = parent_map.get(node)
            container = parent_map.get(anchor) if anchor is not None else None
            if container is None:
                continue
            run = create_text_run(field["placeholder"], _NS_W)
            container.insert(list(container).index(anchor) + 1, run)
            parent_map[run] = container
        return fields

    _RE_INDEX_MATH = re.compile("\x00(\\d+)\x00")

    @classmethod
    def _index_level(
        cls, level: str, eq_latex: Dict[int, str], plain: Dict[int, str],
        for_sort: bool,
    ) -> str:
        """One level of an index term, with its math as $LaTeX$ or plain text."""
        out = []
        for k, piece in enumerate(cls._RE_INDEX_MATH.split(level)):
            if k % 2 == 0:
                # Text. "|" separates levels in a marker and "@" sort key from
                # display; brackets would end the marker; "*" would start
                # emphasis. makeindex quoting is left to the renderer, after
                # the Markdown is parsed.
                out.append(piece.replace("|", "/").replace("@", " at ")
                           .replace("[", "(").replace("]", ")").replace("*", "\\*"))
                continue
            idx = int(piece)
            if for_sort:
                out.append(plain.get(idx, "").replace("|", "/").replace("@", " at "))
                continue
            latex = (eq_latex.get(idx) or "").strip()
            if not latex:
                continue
            if cls._is_text_letter(latex):
                out.append(latex)
                continue
            latex = (latex.replace("[", "\\lbrack ").replace("]", "\\rbrack ")
                     .replace("|", "\\vert "))
            out.append("$" + latex + "$")
        return re.sub(r"\s+", " ", "".join(out)).strip()

    @classmethod
    def _index_marker(cls, field: Dict[str, Any], eq_latex: Dict[int, str]) -> str:
        r"""``XE "main:sub"`` -> ``[index:main|sub]``; math levels get sort@display."""
        code = "".join(
            p if isinstance(p, str) else f"\x00{p[1]}\x00" for p in field["parts"]
        )
        # Straight or curly quotes: a Word typo opens a term with “
        m = re.match(r'\s*XE\s+["\u201c\u201d](.*?)["\u201c\u201d](?=\s|\\|$)', code, re.S)
        if not m:
            return ""
        plain = {p[1]: p[2] for p in field["parts"] if not isinstance(p, str)}
        term = m.group(1).replace("\\:", "\x01")
        levels = []
        for level in term.split(":"):
            level = level.replace("\x01", ":").strip()
            if not level:
                continue
            display = cls._index_level(level, eq_latex, plain, for_sort=False)
            if cls._RE_INDEX_MATH.search(level):
                sort = cls._index_level(level, eq_latex, plain, for_sort=True)
                if sort and display:
                    display = f"{sort}@{display}"
            if display:
                levels.append(display)
        if not levels:
            return ""
        return "[index:" + "|".join(levels) + "]"

    def _splice_index(self, markdown: str, eq_latex: Dict[int, str]) -> str:
        """Put the index markers in place and drop Word's printed index."""
        if not self._index_fields:
            return markdown
        for field in self._index_fields:
            markdown = markdown.replace(
                field["placeholder"], self._index_marker(field, eq_latex)
            )

        lines = markdown.split("\n")
        # A marker in a heading moves to the start of the section's first
        # paragraph: \index inside a sectioning command is fragile
        pending: List[str] = []
        for i, line in enumerate(lines):
            if line.startswith("#"):
                found = re.findall(r"\[index:[^\]]*\]", line)
                if found:
                    pending += found
                    lines[i] = re.sub(r"\[index:[^\]]*\]", "", line).rstrip()
            elif (pending and line.strip()
                  and not line.lstrip().startswith(("$$", "```", "!", "|", ">", "\\["))):
                # After a list bullet, never before it
                bullet = re.match(r"\s*(?:[-+*]|\d+[.)])\s+", line)
                cut = bullet.end() if bullet else 0
                lines[i] = line[:cut] + "".join(pending) + line[cut:]
                pending = []
        markdown = "\n".join(lines)

        # "[index:x](" would read as a link: move the markers (a run of them
        # together, never splitting it) past that word
        markdown = re.sub(
            r"((?:\[index:[^\]]*\])+)((?!\[index:)[(\[{]\S*)", r"\2\1", markdown
        )

        # Word's printed index, with Word's page numbers, gives way to the
        # index built from the markers
        markdown = re.sub(
            r"^# +\**Index\**(?:[ \t]+\{[^}\n]*\})?[ \t]*\n.*?(?=^#{1,2} |^\[\^|\Z)", "", markdown,
            count=1, flags=re.M | re.S,
        )
        return markdown

    # ------------------------------------------------------------------
    # Phase 2a: pandoc on structure-only docx
    # ------------------------------------------------------------------

    def _run_pandoc(
        self, docx_path: Path, media_dir: Path, extra_args: List[str]
    ) -> str:
        """Run pandoc on the math-free docx to get structural markdown."""
        return pypandoc.convert_file(
            str(docx_path),
            "markdown",
            format="docx",
            extra_args=extra_args,
        )

    # ------------------------------------------------------------------
    # Phase 2b: batch equation conversion
    # ------------------------------------------------------------------

    def _batch_convert_equations(
        self, equations: List[Dict[str, Any]], tmp_dir: Path
    ) -> Dict[int, str]:
        """Create a batch .docx with one equation per paragraph, convert all at once.

        Returns:
            dict mapping equation idx → LaTeX string
        """
        register_omml_namespaces()

        # Build a minimal document.xml with marker + equation pairs
        ns_w = NAMESPACES["w"]
        ns_m = NAMESPACES["m"]

        doc_ns = {
            "xmlns:w": ns_w,
            "xmlns:m": ns_m,
            "xmlns:r": NAMESPACES["r"],
        }

        # Create document element
        body_xml = self._build_batch_document(equations, ns_w)

        # Create minimal .docx structure
        batch_dir = tmp_dir / "batch_docx"
        batch_dir.mkdir()

        word_dir = batch_dir / "word"
        word_dir.mkdir()

        # Write document.xml
        (word_dir / "document.xml").write_text(body_xml, encoding="utf-8")

        # Write minimal [Content_Types].xml
        (batch_dir / "[Content_Types].xml").write_text(
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
            '<Default Extension="xml" ContentType="application/xml"/>'
            '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
            '<Override PartName="/word/document.xml" '
            'ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
            '</Types>',
            encoding="utf-8",
        )

        # Write _rels/.rels
        rels_dir = batch_dir / "_rels"
        rels_dir.mkdir()
        (rels_dir / ".rels").write_text(
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" '
            'Target="word/document.xml"/>'
            '</Relationships>',
            encoding="utf-8",
        )

        batch_docx = tmp_dir / "batch.docx"
        rezip_docx(batch_dir, batch_docx)

        # Convert with pandoc
        raw_md = pypandoc.convert_file(
            str(batch_docx),
            "markdown",
            format="docx",
            extra_args=["--wrap=none"],
        )

        # Parse output: markers like @@EQ_0042@@ followed by equation LaTeX
        return self._parse_batch_output(raw_md, equations)

    def _build_batch_document(
        self, equations: List[Dict[str, Any]], ns_w: str
    ) -> str:
        """Build a minimal Word document XML containing marker paragraphs
        and equation paragraphs."""
        parts = [
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
            '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
            'xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math" '
            'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">',
            "<w:body>",
        ]

        for eq in equations:
            marker = f"@@EQ_{eq['idx']:04d}@@"
            # Marker paragraph
            parts.append(
                f"<w:p><w:r><w:t xml:space=\"preserve\">{marker}</w:t></w:r></w:p>"
            )
            # Equation paragraph — inject raw XML
            parts.append(f"<w:p>{eq['xml']}</w:p>")

        parts.append("</w:body></w:document>")
        return "\n".join(parts)

    def _parse_batch_output(
        self, raw_md: str, equations: List[Dict[str, Any]]
    ) -> Dict[int, str]:
        """Parse pandoc's markdown output from the batch document.

        Expects alternating marker lines (@@EQ_NNNN@@) and equation content.
        """
        result: Dict[int, str] = {}
        lines = raw_md.split("\n")
        # Any number of digits: {idx:04d} grows past four at equation 10,000
        marker_re = re.compile(r"@@EQ_(\d+)@@")

        i = 0
        while i < len(lines):
            m = marker_re.search(lines[i])
            if m:
                eq_idx = int(m.group(1))
                # Collect lines until next marker or end
                content_lines = []
                i += 1
                while i < len(lines):
                    if marker_re.search(lines[i]):
                        break
                    content_lines.append(lines[i])
                    i += 1

                # Strip leading/trailing blank lines
                while content_lines and not content_lines[0].strip():
                    content_lines.pop(0)
                while content_lines and not content_lines[-1].strip():
                    content_lines.pop()

                latex = "\n".join(content_lines).strip()
                # Strip any delimiters pandoc may have added
                latex = self._strip_delimiters(latex)
                latex = self._clean_latex(latex)
                result[eq_idx] = latex
            else:
                i += 1

        return result

    @staticmethod
    def _strip_delimiters(latex: str) -> str:
        """Remove any ``$``, ``$$``, ``\\(...\\)``, ``\\[...\\]`` wrapping."""
        s = latex.strip()
        # Display: $$...$$ or \[...\]
        if s.startswith("$$") and s.endswith("$$"):
            s = s[2:-2].strip()
        elif s.startswith("\\[") and s.endswith("\\]"):
            s = s[2:-2].strip()
        # Inline: $...$ or \(...\)
        elif s.startswith("$") and s.endswith("$") and not s.startswith("$$"):
            s = s[1:-1].strip()
        elif s.startswith("\\(") and s.endswith("\\)"):
            s = s[2:-2].strip()
        return s

    # ------------------------------------------------------------------
    # Phase 3: splice
    # ------------------------------------------------------------------

    # Regex to extract \tag{...} from equation content
    _RE_TAG = re.compile(r'\s*\\tag\{([^}]+)\}')

    def _splice(
        self,
        markdown: str,
        eq_latex: Dict[int, str],
        equations: List[Dict[str, Any]],
    ) -> str:
        """Replace placeholders in *markdown* with delimited LaTeX."""
        for eq in equations:
            idx = eq["idx"]
            placeholder = eq["placeholder"]
            latex = eq_latex.get(idx, "")

            if not latex:
                # Remove placeholder if equation came back empty
                markdown = markdown.replace(placeholder, "")
                continue

            if eq["kind"] == "display":
                # Detect QED-only equations (just \square or \blacksquare)
                stripped = self._strip_array_wrapper(latex).strip()
                if stripped in (r'\square', r'\blacksquare'):
                    replacement = f"\n\n\\hfill ${stripped}$\n\n"
                elif self._is_wide_equation(latex):
                    # Extract \tag from content — it can't live inside
                    # $\displaystyle...$, must go in the \[...\] wrapper
                    tag_match = self._RE_TAG.search(latex)
                    latex_body = self._RE_TAG.sub("", latex)
                    # Use \sbox to measure, only shrink if wider than
                    # \linewidth (never scale up)
                    if tag_match:
                        tag_str = f"\\tag{{{tag_match.group(1)}}}"
                        replacement = (
                            "\n\n```{=latex}\n"
                            f"\\[\\sbox0{{$\\displaystyle {latex_body}$}}%\n"
                            "\\ifdim\\wd0>\\linewidth"
                            "\\resizebox{\\linewidth}{!}{\\usebox0}"
                            "\\else\\usebox0\\fi"
                            f" {tag_str}\n\\]\n```\n\n"
                        )
                    else:
                        replacement = (
                            "\n\n```{=latex}\n"
                            "\\[\\sbox0{$\\displaystyle\n"
                            f"{latex_body}\n"
                            "$}%\n"
                            "\\ifdim\\wd0>\\linewidth"
                            "\\resizebox{\\linewidth}{!}{\\usebox0}"
                            "\\else\\usebox0\\fi\n"
                            "\\]\n```\n\n"
                        )
                else:
                    replacement = f"\n\n$$\n{latex}\n$$\n\n"
            else:
                # Inline math — strip \tag (equation numbers don't apply)
                latex = self._RE_TAG.sub("", latex)
                if self._is_text_letter(latex):
                    # A letter such as the é of Poincaré typed as an equation
                    replacement = latex
                else:
                    replacement = f"${latex}$"

            markdown = markdown.replace(placeholder, replacement)

        # Fix adjacent inline math: $...$$ → $...$ $ (add space)
        markdown = re.sub(r'\$\$(?!\n)', _fix_adjacent_inline, markdown)

        # Collapse excessive blank lines introduced by display splicing
        markdown = re.sub(r"\n{4,}", "\n\n\n", markdown)

        return markdown

    # ------------------------------------------------------------------
    # Helpers
    # ------------------------------------------------------------------

    # Strip \begin{array}{...}...\end{array} wrapper to get inner content
    _RE_ARRAY_WRAPPER = re.compile(
        r'^\s*\\begin\{array\}\{[^}]*\}\s*(.*?)\s*\\end\{array\}\s*$',
        re.DOTALL,
    )

    @staticmethod
    def _strip_array_wrapper(latex: str) -> str:
        """Remove outer \\begin{array}...\\end{array} wrapper if present."""
        m = MathExtractor._RE_ARRAY_WRAPPER.match(latex)
        return m.group(1) if m else latex

    @staticmethod
    def _is_wide_equation(latex: str) -> bool:
        """Return True if a display equation is likely to overflow the page.

        Heuristics:
        - Longest line exceeds the character threshold, OR
        - Contains multiple matrix environments (pmatrix, bmatrix, etc.)
        """
        longest = max((len(line) for line in latex.split("\n")), default=0)
        if longest > MathExtractor._WIDE_EQ_THRESHOLD:
            return True
        # Multiple matrices on one line (e.g. matrix product chains)
        matrix_count = len(re.findall(
            r"\\begin\{[pbBvV]?matrix\}", latex
        ))
        if matrix_count >= 3:
            return True
        return False

    # Regex for equation numbers like #(1.1.46) or #(1.1.13a)
    # Captures the number so we can convert #(1.1.46) → \tag{1.1.46}
    # Uses re.MULTILINE so $ matches end-of-line (not just end-of-string),
    # catching numbers before \end{array} on the next line.
    _RE_EQ_NUMBER = re.compile(
        r'[,.\s\\]*'           # optional leading punctuation/space/backslash
        r'\\?#\('              # literal #( possibly with backslash escape
        r'([0-9]+(?:\.[0-9]+)*' # dotted number like 1.1.46 (capture group 1)
        r'[a-z]?)'             # optional letter suffix like 13a
        r'\)'                  # closing paren
        r'[\s\\#]*$'           # trailing whitespace/backslash/hash at end of line
        , re.MULTILINE
    )

    # Threshold (chars) above which a display equation gets \resizebox wrapping
    _WIDE_EQ_THRESHOLD = 300

    # Matches bare \right without a valid delimiter.
    # Catches \right at end-of-line AND \right before \tag{...}
    _RE_BARE_RIGHT = re.compile(
        r'\\right(?=\s*(?:\\tag\{|$))', re.MULTILINE
    )

    # Equation number that OMML wrapped in delimiters: #\left( 1.7.280) \right)
    _RE_EQ_NUMBER_LEFT_RIGHT = re.compile(
        r'(\\?#)\s*\\left\(\s*([0-9]+(?:\.[0-9]+)*[a-z]?)\s*\)?\s*\\right\)'
    )

    # Word's number separator left with no number after it, at end of line
    _RE_EMPTY_EQ_NUMBER = re.compile(r'[ \t]*\\#[ \t]*$', re.MULTILINE)

    # QED marker: \#\square or \# \square (hash + tombstone)
    _RE_QED = re.compile(r'\\?#\s*\\square')

    # A math-alphabet group with no nested braces: \mathbb{\in R,\ }
    _RE_FONT_GROUP = re.compile(
        r'\\(mathbb|mathbf|mathcal|mathfrak|mathrm|mathit|mathsf|mathscr)'
        r'\{([^{}]*)\}'
    )
    _RE_FONT_TOKEN = re.compile(r'\\[A-Za-z]+|\\.|\s+|.', re.DOTALL)

    # Control words that are letters and so belong inside the font group
    _FONT_GLYPH_WORDS = frozenset((
        "alpha beta gamma delta epsilon varepsilon zeta eta theta vartheta "
        "iota kappa lambda mu nu xi pi varpi rho varrho sigma varsigma tau "
        "upsilon phi varphi chi psi omega Gamma Delta Theta Lambda Xi Pi "
        "Sigma Upsilon Phi Psi Omega ell imath jmath hbar"
    ).split())

    # Operators, relations and spacing that Word's font runs sweep into the
    # group; they are moved outside it
    _FONT_OPERATOR_WORDS = frozenset((
        "in notin ni subset subseteq supset supseteq rightarrow leftarrow "
        "Rightarrow Leftarrow leftrightarrow to mapsto leq geq le ge neq ne "
        "times cdot bullet circ pm mp forall exists cup cap setminus wedge "
        "vee land lor oplus otimes approx equiv sim cong propto quad qquad "
        "ldots cdots dots"
    ).split())
    _FONT_OPERATOR_CHARS = frozenset("+-=<>,;:.()[]|'!/*")

    # Greek capitals drawn like Latin letters have no TeX command, and the
    # math fonts carry no glyph for the Unicode ones; Word's math font sets
    # them upright
    _GREEK_LATIN_CAPITALS = {
        "Α": "A", "Β": "B", "Ε": "E", "Ζ": "Z", "Η": "H", "Ι": "I", "Κ": "K",
        "Μ": "M", "Ν": "N", "Ο": "O", "Ρ": "P", "Τ": "T", "Χ": "X",
    }
    # amsfonts' \mathbb has capitals only; lowercase needs bbm's \mathbbm
    _RE_LOWER_MATHBB = re.compile(r"\\mathbb\{([a-z]+)\}")

    @classmethod
    def _fix_math_glyphs(cls, content: str) -> str:
        r"""Replace characters that would print as nothing in the PDF.

        - Greek capital lookalikes: ``Κ`` -> ``\mathrm{K}``
        - pandoc's harpoon accent ``\overset{⃑}{E}`` (U+20D1, a combining
          mark) -> ``\overset{\rightharpoonup}{E}``
        - lowercase double-struck ``\mathbb{c}`` -> ``\mathbbm{c}`` (the
          front matter then loads bbm)
        """
        for greek, latin in cls._GREEK_LATIN_CAPITALS.items():
            content = content.replace(greek, f"\\mathrm{{{latin}}}")
        content = content.replace("\\overset{\u20d1}", "\\overset{\\rightharpoonup}")
        content = cls._fix_group_braces(content)
        content = cls._RE_MATHBF_GROUP.sub(cls._bold_greek, content)
        return cls._RE_LOWER_MATHBB.sub(r"\\mathbbm{\1}", content)

    # Capital Greek in \mathbf is taken from the bold text font at its old
    # 7-bit slots (Delta is 1, Omega 10): a Unicode font has only control
    # characters there, and the letter prints as nothing
    _RE_UPPER_GREEK = re.compile(
        r"(\\(?:Gamma|Delta|Theta|Lambda|Xi|Pi|Sigma|Upsilon|Phi|Psi|Omega)(?![A-Za-z]))"
    )
    _RE_MATHBF_GROUP = re.compile(r"\\mathbf\{([^{}]*)\}")

    @classmethod
    def _bold_greek(cls, match: re.Match) -> str:
        r"""``\mathbf{\Delta}`` -> ``\boldsymbol{\Delta}``, letters kept in ``\mathbf``."""
        if not cls._RE_UPPER_GREEK.search(match.group(1)):
            return match.group(0)
        out = []
        for piece in cls._RE_UPPER_GREEK.split(match.group(1)):
            if cls._RE_UPPER_GREEK.fullmatch(piece):
                out.append("\\boldsymbol{" + piece + "}")
            elif piece.strip():
                out.append("\\mathbf{" + piece + "}")
            else:
                out.append(piece)
        return "".join(out)

    # pandoc sets Word's grouping brace (a groupChr) as a character over or
    # under the group: \overset{X}{︸} is X with a brace beneath it
    _GROUP_BRACES = (
        ("\\overset{", "\ufe38", "\\underbrace"),
        ("\\underset{", "\ufe37", "\\overbrace"),
    )

    @classmethod
    def _fix_group_braces(cls, content: str) -> str:
        r"""``\overset{X}{︸}`` -> ``\underbrace{X}``; ``\underset{X}{︷}`` -> ``\overbrace{X}``."""
        for opener, char, command in cls._GROUP_BRACES:
            start = 0
            while True:
                i = content.find(opener, start)
                if i == -1:
                    break
                j = cls._matching_brace(content, i + len(opener) - 1)
                tail = "{" + char + "}"
                if j != -1 and content.startswith(tail, j + 1):
                    body = content[i + len(opener):j]
                    content = (content[:i] + command + "{" + body + "}"
                               + content[j + 1 + len(tail):])
                start = i + 1
        return content

    @staticmethod
    def _matching_brace(text: str, open_pos: int) -> int:
        """Index of the brace closing the one at *open_pos*, or -1."""
        depth = 0
        i = open_pos
        while i < len(text):
            ch = text[i]
            if ch == "\\":
                i += 2
                continue
            if ch == "{":
                depth += 1
            elif ch == "}":
                depth -= 1
                if depth == 0:
                    return i
            i += 1
        return -1

    @staticmethod
    def _is_text_letter(latex: str) -> bool:
        """True for a lone accented Latin letter, which belongs in the text."""
        if len(latex) != 1:
            return False
        import unicodedata
        return (latex.isalpha() and not latex.isascii()
                and "LATIN" in unicodedata.name(latex, ""))

    @classmethod
    def _split_font_group(cls, match: re.Match) -> str:
        r"""Move operators out of a math-alphabet group, keeping its letters.

        Word applies a font (double-struck, script, fraktur, bold) to a run of
        characters, and the run often takes in the operators beside the
        letter: ``\mathbb{\in R,\ }`` for "∈ ℝ, ". The group is then a single
        ordinary atom, so the relation loses its spacing and the source hides
        what is set in the font. Rewritten: ``\in \mathbb{R},\ ``.

        Groups holding only letters, only operators, or anything unrecognised
        are left unchanged.
        """
        font, body = match.group(1), match.group(2)
        tokens = cls._RE_FONT_TOKEN.findall(body)

        kinds = []
        for tok in tokens:
            if tok.isspace():
                kinds.append("space")
            elif tok.startswith("\\") and tok[1:].isalpha():
                word = tok[1:]
                if word in cls._FONT_GLYPH_WORDS:
                    kinds.append("glyph")
                elif word in cls._FONT_OPERATOR_WORDS:
                    kinds.append("op")
                else:
                    return match.group(0)
            elif tok.startswith("\\"):
                kinds.append("op")  # control symbols: \  \, \; \! \# \{ \} \|
            elif tok.isalnum():
                kinds.append("glyph")
            elif tok in cls._FONT_OPERATOR_CHARS:
                kinds.append("op")
            else:
                return match.group(0)

        if "glyph" not in kinds or "op" not in kinds:
            return match.group(0)

        out = []
        run = []  # glyph tokens (with the spaces between them) awaiting a group
        pending_space = []
        for tok, kind in zip(tokens, kinds):
            if kind == "glyph":
                if run:
                    run.extend(pending_space)
                else:
                    out.extend(pending_space)
                pending_space = []
                run.append(tok)
            elif kind == "space":
                pending_space.append(tok)
            else:
                if run:
                    out.append(f"\\{font}{{{''.join(run)}}}")
                    run = []
                out.extend(pending_space)
                pending_space = []
                out.append(tok)
        if run:
            out.append(f"\\{font}{{{''.join(run)}}}")
        out.extend(pending_space)

        result = "".join(out)
        # A trailing control word must not fuse with a letter that follows
        # the group: \mathfrak{g \times}x -> \mathfrak{g} \times x
        rest = match.string[match.end():match.end() + 1]
        if re.search(r"\\[A-Za-z]+$", result) and rest.isalpha():
            result += " "
        return result

    @staticmethod
    def _clean_latex(content: str) -> str:
        """Post-process a single LaTeX equation string.

        - Strip trailing ``\\ `` (backslash-space), convergence loop
        - Convert equation numbers ``#(1.1.46)`` → ``\\tag{1.1.46}``
        - Fix QED markers ``\\#\\square`` → ``\\square``
        - Fix bare ``\\right`` without delimiter (add invisible ``.``)
        - Fix double subscripts: ``}_{`` → ``{}_{``
        - Fix double superscripts: ``}^{`` → ``{}^{``
        """
        # First, so that spacing moved out of a group is stripped below.
        # Move operators (and a swallowed number separator \#) out of font groups
        content = MathExtractor._RE_FONT_GROUP.sub(
            MathExtractor._split_font_group, content
        )
        content = re.sub(
            r'\\math(?:bb|bf|cal|frak|rm|it|sf|scr)\{([,.;:\s]*\\#)\s*\}', r'\1', content
        )
        content = MathExtractor._RE_EQ_NUMBER_LEFT_RIGHT.sub(r'\1(\2)', content)

        # Strip trailing backslash-space (pandoc artifact)
        limit = 20
        while limit > 0 and content.rstrip().endswith("\\"):
            content = content.rstrip().rstrip("\\").rstrip()
            limit -= 1

        # Extract equation numbers and place \tag at the very end of content
        # so it sits at the outer math level (not inside array/aligned blocks)
        last_number = None
        for m in MathExtractor._RE_EQ_NUMBER.finditer(content):
            last_number = m.group(1)
        content = MathExtractor._RE_EQ_NUMBER.sub("", content)
        # A separator with no number after it: drop it
        content = MathExtractor._RE_EMPTY_EQ_NUMBER.sub("", content)
        # ...and any line it leaves empty: a blank line ends display math
        content = re.sub(r"\n[ \t]*(?=\n)", "", content)
        if last_number:
            content = content.rstrip() + f" \\tag{{{last_number}}}"

        # Fix QED markers: \#\square → \square
        content = MathExtractor._RE_QED.sub(r'\\square', content)

        # Fix bare \right without delimiter (piecewise functions)
        # Pandoc sometimes omits the invisible . delimiter
        content = MathExtractor._RE_BARE_RIGHT.sub(r'\\right.', content)

        # Double subscript/superscript fix
        content = content.replace("}_{", "}{}_{")
        content = content.replace("}^{", "}{}^{")

        # Characters the LaTeX math fonts cannot set (after the fix above,
        # which would read the \mathrm{K}_{m} it produces as a double subscript)
        content = MathExtractor._fix_math_glyphs(content)

        return content.strip()


def _fix_adjacent_inline(match: re.Match) -> str:
    """Callback for adjacent-inline-math regex.

    Only fires when ``$$`` appears NOT followed by a newline (which would
    indicate a display-math opener).  Inserts a space: ``$ $``.
    """
    return "$ $"
