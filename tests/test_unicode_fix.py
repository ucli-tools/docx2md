"""
Tests for the UnicodeFixProcessor module.
"""

import unittest

from docx2md.processors.unicode_fix import UnicodeFixProcessor


class TestUnicodeFixProcessor(unittest.TestCase):

    def _proc(self, config=None):
        return UnicodeFixProcessor(config)

    # ------------------------------------------------------------------
    # Always replacements (context-free)
    # ------------------------------------------------------------------

    def test_nbsp_replaced(self):
        content = "Hello\u00a0World"
        result = self._proc().process(content)
        self.assertEqual(result, "Hello World")

    def test_colon_equals_replaced(self):
        content = "Define $f ≔ x^2$."
        result = self._proc().process(content)
        self.assertIn(":=", result)
        self.assertNotIn("≔", result)

    # ------------------------------------------------------------------
    # Text context replacements
    # ------------------------------------------------------------------

    def test_ell_subscript_in_text(self):
        content = "The norm ℓ₁ is important."
        result = self._proc().process(content)
        self.assertIn(r"$\ell_{1}$", result)
        self.assertNotIn("ℓ₁", result)

    def test_lone_ell_in_text(self):
        content = "The ℓ norm is special."
        result = self._proc().process(content)
        self.assertIn(r"$\ell$", result)
        self.assertNotIn("ℓ", result)

    def test_subscript_digit_in_text(self):
        content = "Variable x₂ appears here."
        result = self._proc().process(content)
        self.assertIn("$_{2}$", result)
        self.assertNotIn("₂", result)

    def test_greek_beta_in_text(self):
        content = "The Greek Β in text."
        result = self._proc().process(content)
        self.assertIn("B", result)
        self.assertNotIn("Β", result)

    # ------------------------------------------------------------------
    # Math context replacements
    # ------------------------------------------------------------------

    def test_ell_in_math(self):
        content = "The norm $ℓ_p$ is defined as..."
        result = self._proc().process(content)
        # ℓ inside math should become \ell
        self.assertIn(r"\ell", result)
        self.assertNotIn("ℓ", result)

    def test_subscript_digit_in_math(self):
        content = "Formula: $x₁ + x₂ = y$"
        result = self._proc().process(content)
        self.assertIn("_1", result)
        self.assertIn("_2", result)
        self.assertNotIn("₁", result)
        self.assertNotIn("₂", result)

    # ------------------------------------------------------------------
    # Ell in math vs text (not double-converted)
    # ------------------------------------------------------------------

    def test_ell_not_double_converted_in_math(self):
        """ℓ inside $...$ should become \ell (not $\ell$)."""
        content = "$ℓ_1$"
        result = self._proc().process(content)
        # Should be $\ell _1$ (or similar), NOT $$\ell$_1$
        self.assertNotIn("$$", result)
        self.assertIn(r"\ell", result)

    # ------------------------------------------------------------------
    # Custom replacements
    # ------------------------------------------------------------------

    def test_custom_always_replacement(self):
        config = {
            "unicode_fix": {
                "custom_replacements": [
                    {"char": "α", "always": r"\alpha"}
                ]
            }
        }
        content = "The angle α is shown."
        result = self._proc(config).process(content)
        self.assertIn(r"\alpha", result)
        self.assertNotIn("α", result)

    # ------------------------------------------------------------------
    # Disabled processor
    # ------------------------------------------------------------------

    def test_disabled(self):
        config = {"unicode_fix": {"enabled": False}}
        content = "Hello\u00a0World ℓ"
        result = self._proc(config).process(content)
        self.assertEqual(result, content)


if __name__ == "__main__":
    unittest.main()


def test_math_alphanumeric_letters_in_text_become_math():
    from docx2md.processors.unicode_fix import UnicodeFixProcessor
    p = UnicodeFixProcessor({})
    assert p.process("If no integer 𝑛 establishes it") == "If no integer $n$ establishes it"
    assert p.process("the algebra 𝔤 and 𝐀") == "the algebra $\\mathfrak{g}$ and $\\mathbf{A}$"
    assert p.process("pair 𝑥𝑦 here") == "pair $xy$ here"
