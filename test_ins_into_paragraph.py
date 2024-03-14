from docx import Document
import docx_tools


def test_insert1():
    doc = Document()
    doc.add_paragraph("")
    assert docx_tools.ins_into_paragraph("FOO", 0, doc.paragraphs[0]).text == "FOO"


def test_insert2():
    doc = Document()
    doc.add_paragraph("FOO")
    assert docx_tools.ins_into_paragraph("BAR", 0, doc.paragraphs[0]).text == "BARFOO"
