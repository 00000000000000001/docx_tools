from docx_tools import cp
from docx import Document

def test_cp():
    document = Document()

    p_src = document.add_paragraph("")

    p_src.add_run("abc").bold = True
    p_src.add_run("defg").underline = True
    p_src.add_run("hijkl").italic = True
    p_dest = document.add_paragraph("")

    cp(5, 7, p_src, p_dest)
    assert p_dest.text == "fgh"
    p_dest._p.clear()

    cp(0, 0, p_src, p_dest)
    assert p_dest.text == "a"
    p_dest._p.clear()

    cp(0, 2, p_src, p_dest)
    assert p_dest.text == "abc"
    p_dest._p.clear()

    cp(0, 3, p_src, p_dest)
    assert p_dest.text == "abcd"
    p_dest._p.clear()

    cp(0, 11, p_src, p_dest)
    assert p_dest.text == "abcdefghijkl"
    p_dest._p.clear()

    cp(7, 11, p_src, p_dest)
    assert p_dest.text == "hijkl"
    p_dest._p.clear()

    cp(11, 11, p_src, p_dest)
    assert p_dest.text == "l"
    p_dest._p.clear()

    cp(0, 11, p_src, p_dest)