"""Tests for docx2md.processors.colours: a Word table's colours kept in the Markdown."""

import io
import zipfile

from docx2md.processors.colours import MIXED, colour_counts, keep_table_colours, md_tables, word_tables

_NS = (
    'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
    'xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"'
)


def _tc(inner: str, fill: str = '') -> str:
    pr = f'<w:tcPr><w:shd w:val="clear" w:color="auto" w:fill="{fill}"/></w:tcPr>' if fill else '<w:tcPr/>'
    return f'<w:tc>{pr}<w:p>{inner}</w:p></w:tc>'


def _run(text: str, colour: str = '') -> str:
    pr = f'<w:rPr><w:color w:val="{colour}"/></w:rPr>' if colour else ''
    return f'<w:r>{pr}<w:t>{text}</w:t></w:r>'


def _math(num: str, den: str, colour: str = '') -> str:
    pr = f'<w:rPr><w:color w:val="{colour}"/></w:rPr>' if colour else ''
    r = lambda t: f'<m:r>{pr}<m:t>{t}</m:t></m:r>'  # noqa: E731
    return f'<m:oMath><m:f><m:num>{r(num)}</m:num><m:den>{r(den)}</m:den></m:f></m:oMath>'


def _tbl(rows) -> str:
    return '<w:tbl><w:tblPr/>' + ''.join('<w:tr>' + ''.join(rows_) + '</w:tr>' for rows_ in rows) + '</w:tbl>'


# Table 1: a header in grey, values in red on light blue, one cell mixing colours.
TABLE_1 = _tbl([
    [_tc(_run('n'), '808080'), _tc(_run('V1'), '808080')],
    [_tc(_run('3'), 'BFBFBF'), _tc(_math('2iπ', '3', 'FF0000'), 'BDD6EE')],
    [_tc(_run('4')), _tc(_run('1', 'FF0000') + _run('2', '0070C0'))],
])
# A layout table holding table 2 (numbers in blue on green), as Word sets side-by-side arrays.
TABLE_2 = _tbl([
    [_tc(_run('1')), _tc(_run('2'))],
    [_tc(_run('2', '0070C0'), 'C5E0B3'), _tc(_run('4'))],
])
LAYOUT = '<w:tbl><w:tblPr/><w:tr><w:tc><w:tcPr/>' + TABLE_2 + '<w:p/></w:tc><w:tc><w:tcPr/><w:p/></w:tc></w:tr></w:tbl>'


def _docx(body: str) -> io.BytesIO:
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, 'w') as z:
        z.writestr('word/document.xml', f'<?xml version="1.0" encoding="UTF-8"?><w:document {_NS}>'
                                        f'<w:body>{body}</w:body></w:document>')
    buf.seek(0)
    return buf


MD = '\n'.join([
    'Some text.',
    '',
    '| **$\\displaystyle n$** | **$\\displaystyle V_{1}$** |',
    '|---|---|',
    '| $\\displaystyle 3$ | $\\displaystyle \\frac{2i\\pi}{3}$ |',
    '| $\\displaystyle 4$ | $\\displaystyle 12$ |',
    '',
    'Between.',
    '',
    '  ---------------',
    '  1         2',
    '  ------- -------',
    '  2         4',
    '  ---------------',
    '',
    'After.',
])


def test_the_word_tables_with_their_colours_a_nested_table_after_its_container():
    tables = word_tables(_docx(TABLE_1 + LAYOUT))
    assert [t.shape for t in tables] == [(3, 2), (1, 2), (2, 2)]
    assert [t.container for t in tables] == [False, True, False]
    first = tables[0]
    assert (first.rows[0][0].fill, first.rows[1][1].fill, first.rows[1][1].colour) == ('808080', 'BDD6EE', 'FF0000')
    assert first.rows[2][1].colour == MIXED and first.rows[1][1].digits == '23'


def test_colours_kept_in_a_pipe_table_and_a_multiline_one():
    out, stats = keep_table_colours(MD, _docx(TABLE_1 + LAYOUT))
    assert stats == {'tables_coloured': 2, 'cells_filled': 5, 'cells_coloured': 2, 'colours_skipped': 1}
    lines = out.split('\n')
    assert lines[2] == ('| `\\cellcolor[HTML]{808080}`{=latex}**$\\displaystyle n$** '
                        '| `\\cellcolor[HTML]{808080}`{=latex}**$\\displaystyle V_{1}$** |')
    assert lines[4] == ('| `\\cellcolor[HTML]{BFBFBF}`{=latex}$\\displaystyle 3$ | `\\cellcolor[HTML]{BDD6EE}`{=latex}'
                        '$\\textcolor[HTML]{FF0000}{\\displaystyle \\frac{2i\\pi}{3}}$ |')
    assert lines[5] == '| $\\displaystyle 4$ | $\\displaystyle 12$ |', 'mixed colours: kept as it was'
    # The multiline table, coloured, is written as a pipe table where it stood.
    assert lines[9:12] == ['| 1 | 2 |', '|---|---|', '| `\\cellcolor[HTML]{C5E0B3}`{=latex}$\\textcolor[HTML]{0070C0}{2}$ | 4 |']
    assert lines[13] == 'After.'
    assert [t.shape for t in md_tables(lines)] == [(3, 2), (2, 2)]


def test_the_check_counts_what_the_word_tables_give():
    docx = _docx(TABLE_1 + LAYOUT)
    out, _ = keep_table_colours(MD, docx)
    want, got, notes = colour_counts(docx, out)
    assert want == got == (5, 2)
    assert notes == ['1 cell(s) mix colours in their text and keep their background only']
    assert colour_counts(docx, MD)[1] == (0, 0), 'before: nothing kept, so the check fails'


def test_a_table_pandoc_did_not_make_is_never_matched_and_keeping_twice_changes_nothing():
    docx = _docx(TABLE_1 + LAYOUT)
    only_second = MD.split('Between.')[1]
    out, stats = keep_table_colours(only_second, docx)
    assert stats['tables_coloured'] == 1 and '`\\cellcolor[HTML]{C5E0B3}`' in out
    once, _ = keep_table_colours(MD, docx)
    assert keep_table_colours(once, docx)[0] == once
