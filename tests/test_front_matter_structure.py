"""
Tests for the FrontMatterStructureProcessor.
"""

import unittest

from docx2md.processors.front_matter_structure import FrontMatterStructureProcessor


def _proc(content, config=None):
    """Helper: run the processor on content and return result."""
    p = FrontMatterStructureProcessor(config or {})
    return p.process(content)


# Simulated YAML block for tests
_YAML = '---\ntitle: "My Book"\nauthor: "Jane Doe"\n---\n'


class TestPageBreakSplitting(unittest.TestCase):

    def test_splits_on_page_break_markers(self):
        content = (
            _YAML
            + '*To my family*\n\n*\\\n*\n\n'
            + 'Copyright 2024 Jane Doe. All rights reserved.\n\n'
            + '# Chapter 1: Introduction\n\nBody.'
        )
        result = _proc(content)
        self.assertIn('# Dedication', result)
        self.assertIn('# Copyright Page', result)

    def test_multiple_page_breaks(self):
        content = (
            _YAML
            + '*For you*\n\n*\\\n*\n\n'
            + 'Some text\n\n*\\\n*\n\n'
            + 'Copyright 2024. ISBN 123.\n\n'
            + '# Chapter 1\n\nBody.'
        )
        result = _proc(content)
        self.assertIn('# Dedication', result)
        self.assertIn('# Copyright Page', result)


class TestDedicationDetection(unittest.TestCase):

    def test_all_italic_paragraphs_detected_as_dedication(self):
        content = _YAML + '*To my beloved family*\n\n# Chapter 1\n\nBody.'
        result = _proc(content)
        self.assertIn('# Dedication', result)
        self.assertIn('*To my beloved family*', result)

    def test_multi_line_italic_dedication(self):
        content = _YAML + '*To my mother*\n*And my father*\n\n# Chapter 1\n\nBody.'
        result = _proc(content)
        self.assertIn('# Dedication', result)

    def test_non_italic_not_dedication(self):
        content = _YAML + 'This is plain text, not italic.\n\n# Chapter 1\n\nBody.'
        result = _proc(content)
        self.assertNotIn('# Dedication', result)


class TestCopyrightDetection(unittest.TestCase):

    def test_copyright_keyword(self):
        content = _YAML + 'Copyright 2024 Jane Doe.\n\n# Chapter 1\n\nBody.'
        result = _proc(content)
        self.assertIn('# Copyright Page', result)

    def test_isbn_keyword(self):
        content = _YAML + 'ISBN 978-0-123456-78-9\n\n# Chapter 1\n\nBody.'
        result = _proc(content)
        self.assertIn('# Copyright Page', result)

    def test_all_rights_reserved(self):
        content = _YAML + 'All rights reserved.\n\n# Chapter 1\n\nBody.'
        result = _proc(content)
        self.assertIn('# Copyright Page', result)

    def test_published_by(self):
        content = _YAML + 'Published by Acme Press.\n\n# Chapter 1\n\nBody.'
        result = _proc(content)
        self.assertIn('# Copyright Page', result)


class TestTitlePageRepeatStripping(unittest.TestCase):

    def test_title_and_author_stripped(self):
        content = _YAML + 'My Book\n\nJane Doe\n\n# Chapter 1\n\nBody.'
        result = _proc(content)
        # The title page repeat should be stripped
        self.assertNotIn('My Book\n\nJane Doe', result)
        self.assertIn('# Chapter 1', result)
        self.assertIn('Body.', result)

    def test_title_alone_not_stripped(self):
        """If only title (no author) is found, don't strip."""
        content = _YAML + 'My Book\n\nSome other content.\n\n# Chapter 1\n\nBody.'
        result = _proc(content)
        # Should remain as unknown, not stripped
        self.assertIn('My Book', result)


class TestPageBreakMarkerRemoval(unittest.TestCase):

    def test_page_break_markers_removed(self):
        content = (
            _YAML
            + '*Dedicated to all*\n\n*\\\n*\n\n'
            + 'Copyright 2024.\n\n'
            + '# Chapter 1\n\nBody.'
        )
        result = _proc(content)
        self.assertNotIn('*\\', result)


class TestPassthrough(unittest.TestCase):

    def test_no_front_matter_passes_through(self):
        content = _YAML + '# Chapter 1\n\nBody text here.'
        result = _proc(content)
        self.assertEqual(content, result)

    def test_empty_pre_heading_passes_through(self):
        content = _YAML + '\n# Chapter 1\n\nBody.'
        result = _proc(content)
        self.assertIn('# Chapter 1', result)

    def test_no_yaml_with_heading(self):
        content = '# Chapter 1\n\nBody text.'
        result = _proc(content)
        self.assertEqual(content, result)

    def test_content_without_headings_or_yaml(self):
        content = 'Just some plain text with no structure.'
        result = _proc(content)
        self.assertEqual(content, result)


class TestCombinedScenario(unittest.TestCase):

    def test_full_front_matter_pipeline(self):
        """Simulate a real .docx conversion with title repeat, dedication, copyright."""
        content = (
            _YAML
            + 'My Book\n\nJane Doe\n\n'  # title page repeat
            + '*\\\n*\n\n'
            + '*To everyone who believed*\n\n'  # dedication
            + '*\\\n*\n\n'
            + 'Copyright 2024 Jane Doe. All rights reserved.\n'
            + 'ISBN 978-0-123456-78-9\n'
            + 'Published by Acme Press.\n\n'  # copyright
            + '# Chapter 1: The Beginning\n\nOnce upon a time.'
        )
        result = _proc(content)
        # Title repeat stripped
        lines_before_ch1 = result.split('# Chapter 1')[0]
        self.assertNotIn('My Book\n\nJane Doe', lines_before_ch1)
        # Dedication present
        self.assertIn('# Dedication', result)
        self.assertIn('*To everyone who believed*', result)
        # Copyright present
        self.assertIn('# Copyright Page', result)
        self.assertIn('ISBN 978-0-123456-78-9', result)
        # Chapter preserved
        self.assertIn('# Chapter 1: The Beginning', result)
        self.assertIn('Once upon a time.', result)


if __name__ == '__main__':
    unittest.main()
