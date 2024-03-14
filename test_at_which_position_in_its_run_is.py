
from docx_tools import at_which_position_in_its_run_is
from docx import Document

def test_all():
    document = Document()

    p_src = document.add_paragraph("")

    p_src.add_run("abc").bold = True
    p_src.add_run("defg").underline = True
    p_src.add_run("hijkl").italic = True

    assert at_which_position_in_its_run_is(0, p_src) == 0
    assert at_which_position_in_its_run_is(1, p_src) == 1
    assert at_which_position_in_its_run_is(2, p_src) == 2

    assert at_which_position_in_its_run_is(3, p_src) == 0
    assert at_which_position_in_its_run_is(4, p_src) == 1
    assert at_which_position_in_its_run_is(5, p_src) == 2
    assert at_which_position_in_its_run_is(6, p_src) == 3

    assert at_which_position_in_its_run_is(7, p_src) == 0
    assert at_which_position_in_its_run_is(8, p_src) == 1
    assert at_which_position_in_its_run_is(9, p_src) == 2
    assert at_which_position_in_its_run_is(10, p_src) == 3
    assert at_which_position_in_its_run_is(11, p_src) == 4

    assert at_which_position_in_its_run_is(12, p_src) == None
    assert at_which_position_in_its_run_is(-1, p_src) == None