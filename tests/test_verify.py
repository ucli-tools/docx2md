"""Tests for docx2md.verify: a conversion is compared with its document."""

import zipfile
from pathlib import Path

import pytest

from docx2md.verify import verify

_NS = (
    'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
    'xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"'
)

_BODY = (
    '<w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr>'
    '<w:r><w:t>Matrix Groups</w:t></w:r></w:p>'
    '<w:p><w:r><w:t xml:space="preserve">Every closure property holds for the group</w:t></w:r>'
    '<w:r><w:footnoteReference w:id="2"/></w:r>'
    '<w:r><w:t xml:space="preserve"> and its elements, since</w:t></w:r></w:p>'
    '<w:p><m:oMathPara><m:oMath><m:r><m:t>x=1.#(1.1.1)</m:t></m:r></m:oMath></m:oMathPara></w:p>'
    '<w:p><w:r><w:t xml:space="preserve">which completes the argument about vectors</w:t></w:r>'
    '<w:r><w:fldChar w:fldCharType="begin"/></w:r>'
    '<w:r><w:instrText xml:space="preserve"> XE "vector" </w:instrText></w:r>'
    '<w:r><w:fldChar w:fldCharType="end"/></w:r>'
    '<w:r><w:t>.</w:t></w:r></w:p>'
)

_MD = (
    '---\ntitle: "T"\n---\n\n'
    '# Matrix Groups\n\n'
    'Every closure property holds for the group[^1] and its elements, since\n\n'
    '$$\nx = 1. \\tag{1.1.1}\n$$\n\n'
    'which completes the argument about vectors[index:vector].\n\n'
    '[^1]: A note.\n'
)


def _docx(tmp_path: Path, body: str = _BODY) -> Path:
    path = tmp_path / 'book.docx'
    with zipfile.ZipFile(path, 'w') as z:
        z.writestr('word/document.xml',
                   f'<?xml version="1.0" encoding="UTF-8"?><w:document {_NS}>'
                   f'<w:body>{body}</w:body></w:document>')
    return path


def _md(tmp_path: Path, text: str) -> Path:
    path = tmp_path / 'book.md'
    path.write_text(text, encoding='utf-8')
    return path


def _failed(checks):
    return {c.name for c in checks if not c.ok}


@pytest.fixture(autouse=True)
def _no_pandoc(monkeypatch):
    # Tables are counted through pandoc's reader; these documents have none
    monkeypatch.setattr('docx2md.verify._count_tables', lambda markdown: 0)


def test_faithful_conversion_passes(tmp_path):
    checks = verify(_docx(tmp_path), _md(tmp_path, _MD))
    assert _failed(checks) == set()


def test_lost_punctuation_before_number_fails(tmp_path):
    md = _MD.replace('x = 1. \\tag', 'x = 1 \\tag')
    checks = verify(_docx(tmp_path), _md(tmp_path, md))
    assert _failed(checks) == {'punctuation before equation numbers'}


def test_punctuation_inside_a_group_counts(tmp_path):
    md = _MD.replace('x = 1. \\tag', 'x = {1.} \\tag')
    checks = verify(_docx(tmp_path), _md(tmp_path, md))
    assert _failed(checks) == set()


def test_lost_footnote_fails(tmp_path):
    md = _MD.replace('group[^1]', 'group').replace('[^1]: A note.\n', '')
    checks = verify(_docx(tmp_path), _md(tmp_path, md))
    assert 'footnotes' in _failed(checks)


def test_footnote_mark_before_a_colon_is_a_mark(tmp_path):
    md = _MD.replace('group[^1] and', 'group[^1]: and')
    checks = verify(_docx(tmp_path), _md(tmp_path, md))
    assert 'footnotes' not in _failed(checks)


def test_word_split_by_index_marker_fails(tmp_path):
    md = _MD.replace('vectors[index:vector]', 'vector[index:vector]s')
    checks = verify(_docx(tmp_path), _md(tmp_path, md))
    assert 'leftovers' in _failed(checks)


def test_lost_equation_number_fails(tmp_path):
    md = _MD.replace(' \\tag{1.1.1}', '')
    checks = verify(_docx(tmp_path), _md(tmp_path, md))
    assert 'equation numbers' in _failed(checks)


def test_display_with_footnote_mark_is_read(tmp_path):
    """A note on a display equation must not unbalance the $ pairs."""
    md = _MD.replace('x = 1. \\tag{1.1.1}\n$$', 'x = 1. \\tag{1.1.1}\n$$[^2]') + '[^2]: Two.\n'
    docx = _docx(tmp_path, _BODY.replace(
        '#(1.1.1)</m:t></m:r>', '#(1.1.1)</m:t></m:r><m:r><w:footnoteReference w:id="3"/></m:r>'))
    checks = verify(docx, _md(tmp_path, md))
    assert _failed(checks) == set()


def test_side_by_side_images_without_caption_fail(tmp_path):
    md = _MD + '\n![](./img/a.png) ![Figure 1.1 -- Two views.](./img/b.png)\n'
    checks = verify(_docx(tmp_path), _md(tmp_path, md))
    assert 'figure captions' in _failed(checks)


def test_limit_moved_off_its_sign_fails(tmp_path):
    md = _MD.replace('x = 1. \\tag', '\\sum_{k = 0}{}^{\\infty} x = 1. \\tag')
    checks = verify(_docx(tmp_path), _md(tmp_path, md))
    assert 'leftovers' in _failed(checks)
