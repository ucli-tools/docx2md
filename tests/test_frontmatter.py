"""
Tests for the generate_yaml_frontmatter module.
"""

import unittest
from pathlib import Path

import yaml

from docx2md.processors.frontmatter import (
    _flip_author_name,
    _insert_newpage_before_sections,
    generate_yaml_frontmatter,
)


def _parse_yaml(fm_str: str) -> dict:
    """Parse the YAML block between --- delimiters."""
    return yaml.safe_load(fm_str.strip("---\n")) or {}


class TestGenerateYamlFrontmatter(unittest.TestCase):

    def _gen(self, doc_props=None, config=None, path=None, content=""):
        return generate_yaml_frontmatter(
            doc_props or {},
            config or {},
            path or Path("test_document.docx"),
            content,
        )

    # ------------------------------------------------------------------
    # Title extraction
    # ------------------------------------------------------------------

    def test_bbm_loaded_for_lowercase_double_struck(self):
        fm, _ = self._gen(content="Text $D_{3\\mathbbm{c}}$ here.")
        self.assertIn("\\usepackage{bbm}", _parse_yaml(fm)["header-includes"])

    def test_figures_kept_in_place_when_document_has_images(self):
        fm, _ = self._gen(content="Text.\n\n![Figure 1](./img/a.png)\n")
        self.assertIn("\\floatplacement{figure}{H}", _parse_yaml(fm)["header-includes"])

    def test_index_enabled_when_markers_present(self):
        fm, _ = self._gen(content="The cone[index:cones] is a surface.")
        self.assertIs(_parse_yaml(fm)["index"], True)

    def test_no_bbm_without_lowercase_double_struck(self):
        fm, _ = self._gen(content="Text $\\mathbb{C}$ here.")
        self.assertNotIn("header-includes", _parse_yaml(fm))

    def test_title_from_doc_properties(self):
        fm_str, _ = self._gen(doc_props={"title": "My Title"})
        fm = _parse_yaml(fm_str)
        self.assertEqual(fm["title"], "My Title")

    def test_title_from_body_bold(self):
        content = "**Introduction to Topology**\n\nBody text."
        fm_str, _ = self._gen(content=content)
        fm = _parse_yaml(fm_str)
        self.assertEqual(fm["title"], "Introduction to Topology")

    def test_title_fallback_from_filename(self):
        fm_str, _ = self._gen(path=Path("my_great_document.docx"), content="Body.")
        fm = _parse_yaml(fm_str)
        self.assertIn("My Great Document", fm["title"])

    # ------------------------------------------------------------------
    # Author and email
    # ------------------------------------------------------------------

    def test_author_from_doc_properties(self):
        fm_str, _ = self._gen(doc_props={"author": "Jane Doe"})
        fm = _parse_yaml(fm_str)
        self.assertEqual(fm["author"], "Jane Doe")

    def test_author_from_body_name_email(self):
        content = "**Title**\n\nJane Doe -- <jane@example.com>\n\nBody."
        fm_str, _ = self._gen(content=content)
        fm = _parse_yaml(fm_str)
        self.assertEqual(fm["author"], "Jane Doe")
        self.assertEqual(fm["email"], "jane@example.com")

    # ------------------------------------------------------------------
    # Author name flipping
    # ------------------------------------------------------------------

    def test_flip_last_first_author(self):
        fm_str, _ = self._gen(doc_props={"author": "Doe, Jane"})
        fm = _parse_yaml(fm_str)
        self.assertEqual(fm["author"], "Jane Doe")

    def test_flip_last_first_middle_author(self):
        fm_str, _ = self._gen(doc_props={"author": "Doe, Jane Marie"})
        fm = _parse_yaml(fm_str)
        self.assertEqual(fm["author"], "Jane Marie Doe")

    def test_no_flip_normal_author(self):
        fm_str, _ = self._gen(doc_props={"author": "Jane Doe"})
        fm = _parse_yaml(fm_str)
        self.assertEqual(fm["author"], "Jane Doe")

    # ------------------------------------------------------------------
    # Subtitle
    # ------------------------------------------------------------------

    def test_subtitle_from_doc_properties(self):
        fm_str, _ = self._gen(doc_props={"subject": "A Subtitle"})
        fm = _parse_yaml(fm_str)
        self.assertEqual(fm["subtitle"], "A Subtitle")

    def test_subtitle_from_body_italic(self):
        content = "**Title**\n\n*An Italic Subtitle*\n\nBody."
        fm_str, _ = self._gen(content=content)
        fm = _parse_yaml(fm_str)
        self.assertEqual(fm["subtitle"], "An Italic Subtitle")

    # ------------------------------------------------------------------
    # mdtexpdf defaults present
    # ------------------------------------------------------------------

    def test_mdtexpdf_defaults_present(self):
        fm_str, _ = self._gen(doc_props={"title": "Test"})
        fm = _parse_yaml(fm_str)
        self.assertIn("format", fm)
        self.assertIn("toc", fm)
        self.assertIn("no_numbers", fm)

    def test_default_format_is_book(self):
        fm_str, _ = self._gen(doc_props={"title": "Test"})
        fm = _parse_yaml(fm_str)
        self.assertEqual(fm["format"], "book")

    def test_active_defaults(self):
        """Verify that key fields are active (uncommented) by default."""
        fm_str, _ = self._gen(doc_props={"title": "Test", "author": "Author"})
        fm = _parse_yaml(fm_str)
        self.assertTrue(fm["toc"])
        self.assertTrue(fm["no_numbers"])
        self.assertTrue(fm["pageof"])
        self.assertTrue(fm["date_footer"])
        self.assertTrue(fm["copyright_page"])
        self.assertTrue(fm["chapters_on_recto"])
        self.assertTrue(fm["drop_caps"])
        self.assertEqual(fm["header_footer_policy"], "all")

    def test_commented_fields_not_in_parsed_yaml(self):
        """Commented-out fields should not appear in parsed YAML."""
        fm_str, _ = self._gen(doc_props={"title": "Test"})
        fm = _parse_yaml(fm_str)
        self.assertNotIn("dedication", fm)
        self.assertNotIn("epigraph", fm)
        self.assertNotIn("cover_image", fm)
        self.assertNotIn("back_cover_image", fm)
        self.assertNotIn("publisher", fm)
        self.assertNotIn("trim_size", fm)
        self.assertNotIn("acknowledgments", fm)

    # ------------------------------------------------------------------
    # Section headers present in raw output
    # ------------------------------------------------------------------

    def test_section_headers_present(self):
        fm_str, _ = self._gen(doc_props={"title": "Test"})
        self.assertIn("# === COMMON METADATA ===", fm_str)
        self.assertIn("# === PDF SETTINGS ===", fm_str)
        self.assertIn("# === PROFESSIONAL BOOK FEATURES ===", fm_str)
        self.assertIn("# === COVER SYSTEM ===", fm_str)
        self.assertIn("# === PRINT FORMAT ===", fm_str)
        self.assertIn("# === ACKNOWLEDGMENTS & ABOUT THE AUTHOR ===", fm_str)
        self.assertIn("# === LATEX PACKAGES", fm_str)

    # ------------------------------------------------------------------
    # Config overrides activate commented fields
    # ------------------------------------------------------------------

    def test_override_activates_dedication(self):
        config = {"yaml_frontmatter": {"mdtexpdf": {"dedication": "To my cat."}}}
        fm_str, _ = self._gen(doc_props={"title": "Test"}, config=config)
        fm = _parse_yaml(fm_str)
        self.assertEqual(fm["dedication"], "To my cat.")

    def test_override_changes_format(self):
        config = {"yaml_frontmatter": {"mdtexpdf": {"format": "article"}}}
        fm_str, _ = self._gen(doc_props={"title": "Test"}, config=config)
        fm = _parse_yaml(fm_str)
        self.assertEqual(fm["format"], "article")

    # ------------------------------------------------------------------
    # Re-run guard (already has frontmatter)
    # ------------------------------------------------------------------

    def test_no_double_prepend(self):
        content = "---\ntitle: Existing\n---\n\nBody."
        fm_str, returned_content = self._gen(content=content)
        self.assertEqual(fm_str, "")
        self.assertEqual(returned_content, content)

    # ------------------------------------------------------------------
    # Title block stripped from body
    # ------------------------------------------------------------------

    def test_title_block_stripped_from_body(self):
        content = "**My Title**\n\nBody paragraph."
        config = {"yaml_frontmatter": {"strip_body_title_block": True}}
        _, updated = self._gen(
            doc_props={"title": "My Title"},
            config=config,
            content=content,
        )
        self.assertNotIn("**My Title**", updated)
        self.assertIn("Body paragraph.", updated)

    # ------------------------------------------------------------------
    # Disabled generator
    # ------------------------------------------------------------------

    def test_disabled(self):
        config = {"yaml_frontmatter": {"enabled": False}}
        fm_str, content = self._gen(
            doc_props={"title": "T"},
            config=config,
            content="Body.",
        )
        self.assertEqual(fm_str, "")
        self.assertEqual(content, "Body.")


class TestFlipAuthorName(unittest.TestCase):

    def test_last_first(self):
        self.assertEqual(_flip_author_name("Doe, Jane"), "Jane Doe")

    def test_last_first_middle(self):
        self.assertEqual(_flip_author_name("Doe, Jane Marie"), "Jane Marie Doe")

    def test_no_comma(self):
        self.assertEqual(_flip_author_name("Jane Doe"), "Jane Doe")

    def test_empty(self):
        self.assertEqual(_flip_author_name(""), "")

    def test_trailing_comma_only(self):
        self.assertEqual(_flip_author_name("Name,"), "Name,")


class TestInsertNewpage(unittest.TestCase):

    def test_newpage_before_references(self):
        content = "Last paragraph.\n\n# References\n\nRef 1."
        result = _insert_newpage_before_sections(content)
        self.assertIn("\\newpage\n\n# References", result)

    def test_newpage_before_index(self):
        content = "Some text.\n\n## Index\n\nA, B, C."
        result = _insert_newpage_before_sections(content)
        self.assertIn("\\newpage\n\n## Index", result)

    def test_newpage_before_bibliography(self):
        content = "End.\n\n# Bibliography\n\n[1] Book."
        result = _insert_newpage_before_sections(content)
        self.assertIn("\\newpage\n\n# Bibliography", result)

    def test_no_newpage_for_normal_heading(self):
        content = "Text.\n\n# Chapter 1\n\nIntro."
        result = _insert_newpage_before_sections(content)
        self.assertNotIn("\\newpage", result)

    def test_no_double_newpage(self):
        content = "Text.\n\n# References\n\nRefs."
        result = _insert_newpage_before_sections(content)
        self.assertEqual(result.count("\\newpage"), 1)


if __name__ == "__main__":
    unittest.main()


class TestDocxFonts(unittest.TestCase):
    """The Word document's typefaces carried into the front matter."""

    def _docx(self, tmp: Path) -> Path:
        import zipfile
        path = tmp / "fonts.docx"
        w = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
        with zipfile.ZipFile(path, "w") as z:
            z.writestr("word/document.xml",
                       f'<w:document {w}><w:body>'
                       '<w:r><w:rPr><w:rFonts w:ascii="Cambria"/></w:rPr></w:r>'
                       '<w:r><w:rPr><w:rFonts w:ascii="Cambria"/></w:rPr></w:r>'
                       '<w:r><w:rPr><w:rFonts w:ascii="Cambria Math"/></w:rPr></w:r>'
                       '<w:r><w:rPr><w:rFonts w:ascii="Cambria Math"/></w:rPr></w:r>'
                       '<w:r><w:rPr><w:rFonts w:ascii="Cambria Math"/></w:rPr></w:r>'
                       '</w:body></w:document>')
            z.writestr("word/styles.xml",
                       f'<w:styles {w}><w:style w:styleId="Heading1"><w:rPr>'
                       '<w:rFonts w:asciiTheme="majorHAnsi"/></w:rPr></w:style></w:styles>')
            z.writestr("word/theme/theme1.xml",
                       '<a:theme><a:majorFont><a:latin typeface="Calibri Light"/></a:majorFont>'
                       '<a:minorFont><a:latin typeface="Calibri"/></a:minorFont></a:theme>')
        return path

    def test_body_and_heading_fonts(self):
        import tempfile
        from docx2md.processors.frontmatter import _docx_fonts
        with tempfile.TemporaryDirectory() as d:
            self.assertEqual(_docx_fonts(self._docx(Path(d))), ("Cambria", "Calibri Light"))

    def test_metric_twin_when_word_font_missing(self):
        from docx2md.processors.frontmatter import _resolve_font
        self.assertEqual(_resolve_font("Cambria", {"Caladea"}), "Caladea")
        self.assertEqual(_resolve_font("Cambria", {"Cambria", "Caladea"}), "Cambria")
        self.assertEqual(_resolve_font("Cambria", set()), "")

    def test_front_matter_fonts(self):
        import tempfile
        from unittest import mock
        from docx2md.processors import frontmatter
        with tempfile.TemporaryDirectory() as d, mock.patch.object(
                frontmatter, "_installed_fonts", return_value={"Caladea", "Carlito"}):
            fm, _ = generate_yaml_frontmatter({}, {}, self._docx(Path(d)), "Text.")
        parsed = _parse_yaml(fm)
        self.assertEqual(parsed["mainfont"], "Caladea")
        self.assertEqual(parsed["sansfont"], "Carlito")
        self.assertIs(parsed["headings_sans"], True)

    def test_greek_fallback_when_main_font_lacks_greek(self):
        from unittest import mock
        from docx2md.processors import frontmatter
        covers = lambda font, chars: font != "Caladea"
        with mock.patch.object(frontmatter, "_font_covers", side_effect=covers):
            self.assertEqual(frontmatter._greek_fallback("Caladea", "A Ἰσόν phrase."), "Noto Serif")
            self.assertEqual(frontmatter._greek_fallback("Caladea", "Only $\\sigma$ and $σ$."), "")
            self.assertEqual(frontmatter._greek_fallback("Noto Serif", "A Ἰσόν phrase."), "")


class TestTitleAccents(unittest.TestCase):

    def test_title_takes_the_texts_accents(self):
        from docx2md.processors.frontmatter import _accented_like_text
        text = "The Café series. Café again, and the Cafe Foundation."
        self.assertEqual(_accented_like_text("Cafe", text), "Café")

    def test_title_kept_when_text_has_no_accent(self):
        from docx2md.processors.frontmatter import _accented_like_text
        self.assertEqual(_accented_like_text("Cafe", "The Cafe Foundation."), "Cafe")

    def test_generated_title_accented(self):
        fm, _ = generate_yaml_frontmatter(
            {"title": "Cafe"}, {}, Path("book.docx"), "Café is the study. Café.")
        self.assertEqual(_parse_yaml(fm)["title"], "Café")
