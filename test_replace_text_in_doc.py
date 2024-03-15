from docx import Document
import docx_tools

def test_replace_text_in_doc1():
    doc = Document()
    doc.add_paragraph("")
    docx_tools.replace_text_in_doc(0, 0, 0, 0, doc, "FOO")
    assert docx_tools.doc_outer_text(doc) == "FOO"


def test_replace_text_in_doc2():
    doc = Document()
    doc.add_paragraph("")
    docx_tools.replace_text_in_doc(0, 0, 0, 0, doc, "BAR")
    assert docx_tools.doc_outer_text(doc) == "BAR"


def test_replace_text_in_doc3():
    doc = Document()
    doc.add_paragraph("FOO")
    docx_tools.replace_text_in_doc(0, 0, 2, 0, doc, "BAR")
    assert docx_tools.doc_outer_text(doc) == "BAR"


def test_replace_text_in_doc4():
    doc = Document()
    doc.add_paragraph("FOO")
    doc.add_paragraph("BAR")
    docx_tools.replace_text_in_doc(1, 0, 2, 1, doc, "Hello")
    assert docx_tools.doc_outer_text(doc) == "FHello"