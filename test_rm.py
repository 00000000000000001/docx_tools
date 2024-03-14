from docx_tools import cp, rm
from docx import Document

def test_rm():
    document = Document()

    p_src = document.add_paragraph("")

    p_src.add_run("abc").bold = True
    p_src.add_run("defg").underline = True
    p_src.add_run("hijkl").italic = True
    p_dest = document.add_paragraph("")

    cp(0, 11, p_src, p_dest)
    rm(0, 0, p_dest)
    assert p_dest.text == "bcdefghijkl"

    p_dest._p.clear()
    cp(0, 11, p_src, p_dest)
    rm(5, 7, p_dest)
    assert p_dest.text == "abcdeijkl"

    p_dest._p.clear()
    cp(0, 11, p_src, p_dest)
    rm(0, 2, p_dest)
    assert p_dest.text == "defghijkl"

    p_dest._p.clear()
    cp(0, 11, p_src, p_dest)
    rm(0, 11, p_dest)
    assert p_dest.text == ""

    p_dest._p.clear()
    cp(0, 11, p_src, p_dest)
    rm(1, 1, p_dest)
    assert p_dest.text == "acdefghijkl"

    p_dest._p.clear()
    cp(0, 11, p_src, p_dest)
    rm(1, 12, p_dest)
    assert p_dest.text == "abcdefghijkl"

    p_dest._p.clear()
    cp(0, 11, p_src, p_dest)
    rm(0, 2, p_dest)
    assert p_dest.text == "defghijkl"

    p_dest._p.clear()
    cp(0, 11, p_src, p_dest)
    rm(3, 6, p_dest)
    assert p_dest.text == "abchijkl"