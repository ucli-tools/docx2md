"""
Tests for the WordCleanupProcessor module.
"""

import unittest
from pathlib import Path

from docx2md.processors.cleanup import WordCleanupProcessor


class TestWordCleanupProcessor(unittest.TestCase):

    def _proc(self, config=None, output_dir=None):
        return WordCleanupProcessor(config, output_dir=output_dir or Path('.'))

    # ------------------------------------------------------------------
    # TOC removal
    # ------------------------------------------------------------------

    def test_remove_toc_basic(self):
        content = (
            "# Table of Contents\n\n"
            "- [Introduction](#introduction)\n"
            "- [Chapter 1](#chapter-1)\n\n"
            "# Introduction\n\nReal content."
        )
        result = self._proc().process(content)
        self.assertNotIn("Table of Contents", result)
        self.assertIn("# Introduction", result)
        self.assertIn("Real content.", result)

    def test_remove_toc_with_class(self):
        content = (
            "# Table of Contents {.TOC-Heading}\n\n"
            "- [Intro](#intro)\n\n"
            "# Intro\n\nText."
        )
        result = self._proc().process(content)
        self.assertNotIn("Table of Contents", result)
        self.assertIn("# Intro", result)

    def test_no_toc_unchanged(self):
        content = "# Introduction\n\nSome text without any TOC."
        result = self._proc().process(content)
        self.assertIn("# Introduction", result)
        self.assertIn("Some text without any TOC.", result)

    # ------------------------------------------------------------------
    # Heading markup stripping
    # ------------------------------------------------------------------

    def test_strip_bold_italic_heading(self):
        content = "# ***Title with bold-italic***\n\nBody text."
        result = self._proc().process(content)
        self.assertIn("# Title with bold-italic", result)
        self.assertNotIn("***", result.split('\n')[0])

    def test_strip_bold_heading(self):
        content = "## **Section Title**\n\nParagraph."
        result = self._proc().process(content)
        self.assertIn("## Section Title", result)

    def test_heading_with_math_preserved(self):
        content = "# ***Heading*** $f(x) = x^2$\n\nBody."
        result = self._proc().process(content)
        self.assertIn("$f(x) = x^2$", result)
        self.assertNotIn("***", result.split('\n')[0])

    # ------------------------------------------------------------------
    # Heading ID stripping
    # ------------------------------------------------------------------

    def test_strip_heading_id(self):
        content = "# Introduction {#introduction}\n\nText."
        result = self._proc().process(content)
        self.assertNotIn("{#introduction}", result)

    def test_keep_unnumbered_class(self):
        content = "# Preface {#preface .unnumbered}\n\nText."
        result = self._proc().process(content)
        self.assertIn("{.unnumbered}", result)
        self.assertNotIn("#preface", result)

    def test_strip_id_keep_only_unnumbered(self):
        content = "# Chapter {#chap .TOC-Heading .unnumbered}\n\nText."
        result = self._proc().process(content)
        self.assertIn("{.unnumbered}", result)
        self.assertNotIn("#chap", result)

    # ------------------------------------------------------------------
    # Image size attribute removal
    # ------------------------------------------------------------------

    def test_remove_image_size_attrs(self):
        content = '![alt text](img/image.png){width="3in" height="2in"}\n'
        result = self._proc().process(content)
        self.assertIn("![alt text](img/image.png)", result)
        self.assertNotIn('width=', result)
        self.assertNotIn('height=', result)

    def test_image_without_attrs_unchanged(self):
        content = "![diagram](img/diagram.png)\n"
        result = self._proc().process(content)
        self.assertEqual(content, result)

    # ------------------------------------------------------------------
    # Image path fixing
    # ------------------------------------------------------------------

    def test_fix_absolute_image_path(self):
        output_dir = Path('/tmp/output')
        content = "![fig](/tmp/output/img/image.png)"
        result = self._proc(output_dir=output_dir).process(content)
        self.assertNotIn('/tmp/output/img/image.png', result)
        self.assertIn('img/image.png', result)

    def test_relative_path_unchanged(self):
        content = "![fig](img/image.png)"
        result = self._proc().process(content)
        self.assertIn("img/image.png", result)

    # ------------------------------------------------------------------
    # Config flags
    # ------------------------------------------------------------------

    def test_disable_remove_toc(self):
        config = {"cleanup": {"remove_toc": False}}
        content = "# Table of Contents\n\n- [Intro](#intro)\n\n# Intro\n"
        result = self._proc(config).process(content)
        self.assertIn("Table of Contents", result)

    def test_disable_strip_heading_markup(self):
        config = {"cleanup": {"strip_heading_markup": False}}
        content = "# ***Bold Title***\n\nText."
        result = self._proc(config).process(content)
        self.assertIn("***Bold Title***", result)


if __name__ == "__main__":
    unittest.main()


class TestAutolinkBareUrls:
    """Bare web addresses become autolinks that can break across a line."""

    def test_bare_url_wrapped_trailing_punctuation_kept_outside(self):
        from docx2md.processors.cleanup import _autolink_bare_urls
        assert _autolink_bare_urls("at http://www.geogebra.com, which") == \
            "at <http://www.geogebra.com>, which"

    def test_escapes_removed_inside_autolink(self):
        from docx2md.processors.cleanup import _autolink_bare_urls
        assert _autolink_bare_urls("see http://x.edu/\\~haber/a.pdf*") == \
            "see <http://x.edu/~haber/a.pdf>*"

    def test_existing_links_untouched(self):
        from docx2md.processors.cleanup import _autolink_bare_urls
        text = "a [site](https://a.b/c) and <https://x.y> and ![i](./img/a.png)"
        assert _autolink_bare_urls(text) == text
