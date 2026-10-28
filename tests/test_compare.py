"""Tests for docx2md.compare: two Word documents compared, and one corrected in one place."""

import io
import zipfile
from pathlib import Path

import pytest
from click.testing import CliRunner

from docx2md.cli import main
from docx2md.compare import compare, edit, edit_superscript, locate, locate_all

_NS = (
    'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
    'xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"'
)
_STYLES = (
    f'<?xml version="1.0" encoding="UTF-8"?><w:styles {_NS}>'
    '<w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/>'
    '<w:pPr><w:outlineLvl w:val="0"/></w:pPr></w:style>'
    '<w:style w:type="paragraph" w:styleId="Heading2"><w:name w:val="heading 2"/>'
    '<w:basedOn w:val="Heading1"/><w:pPr><w:outlineLvl w:val="1"/></w:pPr></w:style>'
    '</w:styles>'
)


def _frac(num: str, den: str, colour: str = 'FF0000') -> str:
    def run(t: str) -> str:
        return f'<m:r><w:rPr><w:color w:val="{colour}"/></w:rPr><m:t>{t}</m:t></m:r>'
    return f'<m:oMath><m:f><m:num>{run(num)}</m:num><m:den>{run(den)}</m:den></m:f></m:oMath>'


def _cell(inner: str) -> str:
    return f'<w:tc><w:tcPr><w:tcW w:w="500"/></w:tcPr><w:p>{inner}</w:p></w:tc>'


def _body(v10: str = '10iπ', colour: str = 'FF0000', extra: str = '') -> str:
    return (
        '<w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t>2.1 - Polygons</w:t></w:r></w:p>'
        '<w:p><w:pPr><w:pStyle w:val="Heading2"/></w:pPr><w:r><w:t>Polygons</w:t></w:r></w:p>'
        '<w:p w:rsidR="00AB12CD"><w:r><w:t xml:space="preserve">'
        'The vertices of the first polygons are shown in the table below.</w:t></w:r></w:p>'
        + extra +
        '<w:tbl><w:tblPr/>'
        '<w:tr>' + _cell('<w:r><w:t>n</w:t></w:r>') + _cell('<w:r><w:t>V0</w:t></w:r>')
        + _cell('<w:r><w:t>V10</w:t></w:r>') + '</w:tr>'
        '<w:tr w:rsidR="00FF0011">' + _cell('<w:r><w:t>13</w:t></w:r>') + _cell('<m:oMath><m:r><m:t>ϵ</m:t></m:r></m:oMath>')
        + _cell(_frac(v10, '13', colour)) + '</w:tr>'
        '</w:tbl>'
        '<w:p><w:r><w:t>Table 2.1 – The vertices of the polygons.</w:t></w:r></w:p>'
    )


def _docx(body: str, image: bytes = b'PNG') -> bytes:
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, 'w', zipfile.ZIP_DEFLATED) as z:
        z.writestr('word/document.xml',
                   f'<?xml version="1.0" encoding="UTF-8"?><w:document {_NS}><w:body>{body}</w:body></w:document>')
        z.writestr('word/styles.xml', _STYLES)
        z.writestr('word/media/image1.png', image)
    return buf.getvalue()


def test_the_same_document_has_no_change():
    assert compare(_docx(_body()), _docx(_body())).changes == []


def test_a_value_in_a_table_cell_is_one_change_with_its_place_and_whole_row():
    cmp = compare(_docx(_body()), _docx(_body('20iπ')))
    assert cmp.parts == []
    [c] = cmp.changes
    assert c.kind == 'changed'
    assert c.where == '2.1 - Polygons › Polygons › table 1 of the document (Table 2.1), row 2'
    assert c.before == '13 | ϵ | 10iπ/13'
    assert c.after == '13 | ϵ | 20iπ/13'
    assert c.details == ['cell 3: « 10iπ/13 » → « 20iπ/13 »']


def test_a_colour_alone_is_a_change():
    [c] = compare(_docx(_body()), _docx(_body(colour='1E48FF'))).changes
    assert c.before == c.after
    assert 'formatting changed' in c.details[-1]


def test_what_word_rewrites_on_save_is_not_a_change():
    saved = _body().replace('w:rsidR="00AB12CD"', 'w:rsidR="00999999"').replace('w:rsidR="00FF0011"', 'w:rsidR="00123456"')
    assert compare(_docx(_body()), _docx(saved)).changes == []


def test_a_paragraph_added_and_an_image_changed():
    extra = '<w:p><w:r><w:t>A new sentence.</w:t></w:r></w:p>'
    cmp = compare(_docx(_body()), _docx(_body(extra=extra), image=b'PNG2'))
    assert [(c.kind, c.after) for c in cmp.changes] == [('added', 'A new sentence.')]
    assert cmp.parts == [('word/media/image1.png', 'changed')]


def test_maths_read_as_a_line_of_text():
    body = ('<w:p><m:oMath><m:sSup><m:e><m:r><m:t>e</m:t></m:r></m:e><m:sup>' + _frac('2iπ', '3')[9:-10] +
            '</m:sup></m:sSup><m:d><m:e><m:r><m:t>x+1</m:t></m:r></m:e></m:d></m:oMath></w:p>')
    cmp = compare(_docx(body), _docx(''))
    assert cmp.changes[0].before == 'e^(2iπ/3)(x+1)'


def test_edit_changes_one_value_and_nothing_else():
    old = _docx(_body())
    out = io.BytesIO()
    edit(old, out, '10iπ', '20iπ', table=1, row=2, cell=3)
    new = out.getvalue()
    [c] = compare(old, new).changes
    assert c.details == ['cell 3: « 10iπ/13 » → « 20iπ/13 »']
    with zipfile.ZipFile(io.BytesIO(old)) as a, zipfile.ZipFile(io.BytesIO(new)) as b:
        assert a.namelist() == b.namelist()
        for name in a.namelist():
            if name != 'word/document.xml':
                assert a.read(name) == b.read(name)
        xa, xb = a.read('word/document.xml').decode(), b.read('word/document.xml').decode()
        assert xb == xa.replace('<m:t>10iπ</m:t>', '<m:t>20iπ</m:t>')


def test_edit_refuses_a_text_that_is_not_there_exactly_once():
    with pytest.raises(ValueError, match='0 times'):
        edit(_docx(_body()), io.BytesIO(), '11iπ', '20iπ', table=1, row=2, cell=3)
    twice = _body().replace('<m:t>13</m:t>', '<m:t>10iπ</m:t>', 1)
    with pytest.raises(ValueError, match='2 times'):
        edit(_docx(twice), io.BytesIO(), '10iπ', '20iπ', table=1, row=2, cell=3)
    with pytest.raises(ValueError, match='1 tables, not 2'):
        edit(_docx(_body()), io.BytesIO(), '10iπ', '20iπ', table=2, row=2, cell=3)


def test_edit_in_a_paragraph_named_by_its_words():
    old = _docx(_body())
    out = io.BytesIO()
    edit(old, out, 'first polygons', 'first thirteen polygons', paragraph='vertices of the first')
    [c] = compare(old, out.getvalue()).changes
    assert c.after == 'The vertices of the first thirteen polygons are shown in the table below.'
    with pytest.raises(ValueError, match='0 paragraphs'):
        locate(f'<w:document {_NS}><w:body>{_body()}</w:body></w:document>', paragraph='no such words')


def test_a_paragraph_found_after_empty_ones_and_a_space_kept_before_maths():
    # Word writes an empty paragraph as one tag; the paragraphs after it are still each their own.
    body = ('<w:p w:rsidR="00AA0001"/>'
            '<w:p><w:r><w:t>The order is equal to</w:t></w:r><m:oMath><m:r><m:t>1</m:t></m:r></m:oMath>'
            '<w:r><w:t>.</w:t></w:r></w:p><w:p w:rsidR="00AA0002"/>' + _body())
    old = _docx(body)
    out = io.BytesIO()
    edit(old, out, 'equal to', 'equal to ', paragraph='The order is')
    [c] = compare(old, out.getvalue()).changes
    assert c.before.startswith('The order is equal to1') and c.after.startswith('The order is equal to 1')
    with zipfile.ZipFile(io.BytesIO(out.getvalue())) as z:
        assert '<w:t xml:space="preserve">The order is equal to </w:t>' in z.read('word/document.xml').decode()
    edit(old, io.BytesIO(), 'first polygons', 'first thirteen polygons', paragraph='vertices of the first')


def test_edit_the_nth_of_several_identical_paragraphs():
    twin = '<w:p><w:r><w:t>Zorp teh quill sits here.</w:t></w:r></w:p>'
    old = _docx(twin + '<w:p><w:r><w:t>Between.</w:t></w:r></w:p>' + twin)
    with pytest.raises(ValueError, match='2 paragraphs, not exactly one'):
        edit(old, io.BytesIO(), 'teh', 'the', paragraph='Zorp')
    out = io.BytesIO()
    edit(old, out, 'teh', 'the', paragraph='Zorp', paragraph_nth=2)
    with zipfile.ZipFile(io.BytesIO(out.getvalue())) as z:
        xml = z.read('word/document.xml').decode()
    assert xml.count('Zorp teh quill') == 1 and xml.count('Zorp the quill') == 1
    assert xml.index('Zorp teh quill') < xml.index('Zorp the quill')
    for bad in (0, 3):
        with pytest.raises(ValueError, match=f'no paragraph {bad}'):
            edit(old, io.BytesIO(), 'teh', 'the', paragraph='Zorp', paragraph_nth=bad)


def test_edit_the_nth_of_several():
    body = ('<w:p><w:r><w:t xml:space="preserve">Its order is equal to </w:t></w:r><m:oMath><m:r><m:t>2</m:t></m:r></m:oMath>'
            '<w:r><w:t xml:space="preserve"> and its rank equal to</w:t></w:r><m:oMath><m:r><m:t>1</m:t></m:r></m:oMath></w:p>')
    old = _docx(body + _body())
    with pytest.raises(ValueError, match='2 times, not exactly once'):
        edit(old, io.BytesIO(), ' equal to', ' equal to ', paragraph='Its order')
    with pytest.raises(ValueError, match='no occurrence 3'):
        edit(old, io.BytesIO(), ' equal to', ' equal to ', nth=3, paragraph='Its order')
    out = io.BytesIO()
    edit(old, out, ' equal to', ' equal to ', nth=2, paragraph='Its order')
    [c] = compare(old, out.getvalue()).changes
    assert c.after.startswith('Its order is equal to 2 and its rank equal to 1')


def test_edit_an_equations_exponent():
    def power(sup: str) -> str:
        return ('<m:sSup><m:e><m:r><m:t>a</m:t></m:r></m:e><m:sup>' + sup + '</m:sup></m:sSup>')
    half = '<m:f><m:num><m:r><w:rPr><w:sz w:val="18"/></w:rPr><m:t>3</m:t></m:r></m:num><m:den><m:r><m:t>2</m:t></m:r></m:den></m:f>'
    body = ('<w:p><m:oMath><m:r><m:t>k=</m:t></m:r>' + power(half) + '<m:r><m:t>,#(9.9.9)</m:t></m:r>'
            + power(half) + '</m:oMath></w:p>')
    old = _docx(body + _body())
    with pytest.raises(ValueError, match='2 times, not exactly once'):
        edit_superscript(old, io.BytesIO(), '3/2', '3', paragraph='9.9.9')
    with pytest.raises(ValueError, match='0 times'):
        edit_superscript(old, io.BytesIO(), '5/2', '3', paragraph='9.9.9')
    out = io.BytesIO()
    edit_superscript(old, out, '3/2', '3', nth=1, paragraph='9.9.9')
    [c] = compare(old, out.getvalue()).changes
    assert c.before == 'k=a^(3/2),#(9.9.9)a^(3/2)' and c.after == 'k=a^3,#(9.9.9)a^(3/2)'
    with zipfile.ZipFile(io.BytesIO(out.getvalue())) as z:
        assert '<m:sup><m:r><w:rPr><w:sz w:val="18"/></w:rPr><m:t>3</m:t></m:r></m:sup>' in z.read('word/document.xml').decode()


def test_paragraph_nth_with_a_table_is_refused():
    old = _docx(_body())
    with pytest.raises(ValueError, match='not with --table'):
        edit(old, io.BytesIO(), 'n', 'm', table=1, row=1, cell=1, paragraph_nth=1)


def _twins():
    """Four paragraphs holding 'Zorp', in document order: in the body, in a table cell, in a
    text box inside another paragraph, and one that is an OMML m:oMathPara."""
    return (
        '<w:p><w:r><w:t>Zorp teh one.</w:t></w:r></w:p>'
        '<w:tbl><w:tblPr/><w:tr>' + _cell('<w:r><w:t>Zorp teh two.</w:t></w:r>') + '</w:tr></w:tbl>'
        '<w:p><w:r><w:t>Host.</w:t></w:r><w:r><mc:AlternateContent><mc:Choice><w:drawing><w:txbxContent>'
        '<w:p><w:r><w:t>Zorp teh three.</w:t></w:r></w:p></w:txbxContent></w:drawing></mc:Choice></mc:AlternateContent></w:r></w:p>'
        '<w:p><m:oMathPara><m:oMath><m:r><m:t>Zorp teh four.</m:t></m:r></m:oMath></m:oMathPara></w:p>')


def _xml_of(docx: bytes) -> str:
    with zipfile.ZipFile(io.BytesIO(docx)) as z:
        return z.read('word/document.xml').decode()


def test_paragraph_nth_counts_in_document_order_across_cell_text_box_and_maths():
    old = _docx(_twins())
    assert [locate(_xml_of(old), paragraph='Zorp', paragraph_nth=k)[0] for k in (1, 2, 3, 4)] == sorted(
        locate(_xml_of(old), paragraph='Zorp', paragraph_nth=k)[0] for k in (1, 2, 3, 4))
    for k, word in enumerate(('one', 'two', 'three', 'four'), 1):
        out = io.BytesIO()
        edit(old, out, 'teh', 'the', paragraph='Zorp', paragraph_nth=k)
        xml = _xml_of(out.getvalue())
        assert xml.count('Zorp the ') == 1 and xml.count('Zorp teh ') == 3
        assert f'Zorp the {word}.' in xml
    with pytest.raises(ValueError, match='is in 4 paragraphs'):
        edit(old, io.BytesIO(), 'teh', 'the', paragraph='Zorp')


def test_nth_counts_inside_the_chosen_paragraph_only():
    twin = '<w:p><w:r><w:t>Zorp a a</w:t></w:r><w:r><w:t>a</w:t></w:r></w:p>'
    old = _docx(twin + twin)
    out = io.BytesIO()
    edit(old, out, 'a', 'b', paragraph='Zorp', paragraph_nth=2, nth=3)
    xml = _xml_of(out.getvalue())
    first, second = xml.split('</w:p>')[:2]
    assert first.count('<w:t>a</w:t>') == 1 and 'b' not in first.replace('<w:body>', '')
    assert second.count('<w:t>a</w:t>') == 0 and second.count('<w:t>b</w:t>') == 1 and 'Zorp a a' in second


def test_edit_superscript_in_the_nth_paragraph():
    def power(sup: str) -> str:
        return '<m:sSup><m:e><m:r><m:t>a</m:t></m:r></m:e><m:sup><m:r><m:t>' + sup + '</m:t></m:r></m:sup></m:sSup>'
    twin = '<w:p><m:oMath><m:r><m:t>Zorp=</m:t></m:r>' + power('2') + '</m:oMath></w:p>'
    old = _docx(twin + twin)
    with pytest.raises(ValueError, match='2 paragraphs, not exactly one'):
        edit_superscript(old, io.BytesIO(), '2', '3', paragraph='Zorp')
    out = io.BytesIO()
    edit_superscript(old, out, '2', '3', paragraph='Zorp', paragraph_nth=2)
    xml = _xml_of(out.getvalue())
    assert xml.count('<m:t>2</m:t>') == 1 and xml.count('<m:t>3</m:t>') == 1
    assert xml.index('<m:t>2</m:t>') < xml.index('<m:t>3</m:t>')


def test_locate_lists_every_match_in_order_and_exits_by_result():
    here = Path(__file__).parent
    f = here / 'test_compare_locate.docx'
    runner = CliRunner()
    try:
        f.write_bytes(_docx(_twins()))
        hits = locate_all(f, paragraph='Zorp')
        assert [h[0] for h in hits] == [1, 2, 3, 4]
        assert [h[1] for h in hits] == sorted(h[1] for h in hits)
        assert [h[0] for h in locate_all(f, paragraph='Zorp', paragraph_nth=2)] == [2]
        r = runner.invoke(main, ['locate', str(f), '--paragraph', 'Zorp'])
        assert r.exit_code == 0 and r.output.count('K=') == 4 and 'Zorp teh one.' in r.output
        assert runner.invoke(main, ['locate', str(f), '--paragraph', 'Zorp', '--paragraph-nth', '3']).output.count('K=') == 1
        assert runner.invoke(main, ['locate', str(f), '--paragraph', 'Nothing here']).exit_code == 1
    finally:
        f.unlink(missing_ok=True)


def test_the_report_and_the_command():
    # Files named for this test beside it, removed when it ends.
    here = Path(__file__).parent
    old, new, rep = (here / f'test_compare_{n}' for n in ('old.docx', 'new.docx', 'report.md'))
    runner = CliRunner()
    try:
        old.write_bytes(_docx(_body()))
        assert runner.invoke(main, ['edit', str(old), str(new), '--replace', '10iπ', '20iπ',
                                    '--table', '1', '--row', '2', '--cell', '3']).exit_code == 0
        ok = runner.invoke(main, ['diff', str(old), str(new), '--expect', '1', '-o', str(rep)])
        assert ok.exit_code == 0, ok.output
        text = rep.read_text(encoding='utf-8')
        assert '## 1. changed: 2.1 - Polygons › Polygons › table 1 of the document (Table 2.1), row 2' in text
        assert '- cell 3: « 10iπ/13 » → « 20iπ/13 »' in text and 'SHA-256' in text
        assert runner.invoke(main, ['diff', str(old), str(new), '--expect', '0']).exit_code == 1
        assert runner.invoke(main, ['edit', str(old), str(old), '--replace', 'a', 'b', '--paragraph', 'x']).exit_code == 1
    finally:
        for f in (old, new, rep):
            f.unlink(missing_ok=True)
