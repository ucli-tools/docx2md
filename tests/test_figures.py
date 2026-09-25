"""
Tests for the FigureProcessor module.
"""

import unittest

from docx2md.processors.figures import FigureProcessor


class TestFigureProcessor(unittest.TestCase):

    def _proc(self, config=None):
        return FigureProcessor(config)

    # ------------------------------------------------------------------
    # Basic double-caption fix
    # ------------------------------------------------------------------

    def test_replaces_ai_alt_with_real_caption(self):
        content = (
            "![A computer generated image of a graph automatically generated](img/img.png)\n\n"
            "***Figure 1.*** *-- The Lie algebra structure of the polyplex.*\n\n"
            "Body text continues here."
        )
        result = self._proc().process(content)
        self.assertIn("Figure 1.", result)
        self.assertIn("The Lie algebra structure", result)
        self.assertNotIn("automatically generated", result)
        # Caption paragraph removed
        self.assertEqual(result.count("Figure 1."), 1)

    def test_caption_paragraph_removed(self):
        content = (
            "![AI alt text](img/fig.png)\n\n"
            "***Figure 2.*** *-- Some description here.*\n\n"
            "Next paragraph."
        )
        result = self._proc().process(content)
        # Caption block should not appear separately
        self.assertNotIn("***Figure 2.***", result)
        # Image should still be there
        self.assertIn("img/fig.png", result)

    # ------------------------------------------------------------------
    # Caption with math preserved
    # ------------------------------------------------------------------

    def test_caption_preserves_math(self):
        content = (
            "![generic alt](img/eq.png)\n\n"
            "***Figure 3.*** *-- The formula $E = mc^2$ in context.*\n\n"
            "End."
        )
        result = self._proc().process(content)
        self.assertIn("$E = mc^2$", result)
        self.assertNotIn("generic alt", result)

    def test_italic_run_across_math_leaves_no_asterisk(self):
        content = (
            "![generic alt](img/eq.png)\n\n"
            "***Figure 1.1.3 --*** *The* $D_{3}$ *group permutations.*\n\n"
            "***Figure 1.1.4 --*** *The formula $E = mc^2$ in context.*\n\n"
            "End."
        )
        result = self._proc().process(content)
        alt = result.split("](")[0]
        self.assertNotIn("*", alt)
        self.assertIn("The $D_{3}$ group permutations.", alt)

    def test_escaped_asterisk_kept(self):
        content = "![alt](img/a.png)\n\n***Figure 2.*** *-- The star \\* operator.*\n\nEnd."
        result = self._proc().process(content)
        self.assertIn("\\*", result.split("](")[0])

    # ------------------------------------------------------------------
    # No caption following image
    # ------------------------------------------------------------------

    def test_image_without_caption_unchanged(self):
        content = (
            "![diagram](img/diag.png)\n\n"
            "This paragraph is not a caption.\n\n"
            "More text."
        )
        result = self._proc().process(content)
        self.assertIn("![diagram](img/diag.png)", result)
        self.assertIn("This paragraph is not a caption.", result)

    # ------------------------------------------------------------------
    # Multi-line word-wrapped caption
    # ------------------------------------------------------------------

    def test_multiline_caption_collapsed(self):
        content = (
            "![alt](img/long.png)\n\n"
            "***Figure 4.*** *-- This is a long caption that wraps\n"
            "across multiple lines in the source.*\n\n"
            "Footer."
        )
        result = self._proc().process(content)
        self.assertIn("Figure 4.", result)
        self.assertIn("wraps", result)
        self.assertNotIn("alt", result.split('(img/long.png)')[0].split('![')[-1])

    # ------------------------------------------------------------------
    # Disabled processor
    # ------------------------------------------------------------------

    def test_disabled(self):
        config = {"figures": {"enabled": False}}
        content = (
            "![AI alt text](img/fig.png)\n\n"
            "***Figure 1.*** *-- Real caption.*\n\n"
            "Text."
        )
        result = self._proc(config).process(content)
        # Nothing should change
        self.assertIn("AI alt text", result)
        self.assertIn("***Figure 1.***", result)


if __name__ == "__main__":
    unittest.main()
