from docx_tools import mv
from docx import Document

def test_mv():
    document = Document()

    p_src = document.add_paragraph("")

    p_src.add_run("abc").bold = True
    p_src.add_run("defg").underline = True
    p_src.add_run("hijkl").italic = True
    p_dest = document.add_paragraph("")
    p_dest._p.clear()
    mv(0, 0, p_src, p_dest)
    assert p_src.text == "bcdefghijkl"
    assert p_dest.text == "a"

    p_src._p.clear()
    p_src.add_run("abc").bold = True
    p_src.add_run("defg").underline = True
    p_src.add_run("hijkl").italic = True
    p_dest._p.clear()
    mv(0, 11, p_src, p_dest)
    assert p_src.text == ""
    assert p_dest.text == "abcdefghijkl"

    p_src._p.clear()
    p_src.add_run("abc").bold = True
    p_src.add_run("defg").underline = True
    p_src.add_run("hijkl").italic = True
    p_dest._p.clear()
    mv(3, 6, p_src, p_dest)
    assert p_src.text == "abchijkl"
    assert p_dest.text == "defg"