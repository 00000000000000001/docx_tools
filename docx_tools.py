import copy
from docx.text.paragraph import Paragraph
from docx.oxml.shared import OxmlElement


def doc_text(doc):
    fullText = ""
    for p in doc.paragraphs:
        fullText += p.text
    return fullText


def tables_text(doc):
    fullText = ""
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                fullText += cell.text
    return fullText


def duplicate(p):
    p_new = copy.deepcopy(p)
    p._p.addnext(p_new._p)
    return p_new


def delete_paragraph(paragraph):
    p = paragraph._element
    p.getparent().remove(p)
    paragraph._p = paragraph._element = None


def append_paragraph(paragraph, text=None, style=None):
    try:
        """Insert a new paragraph after the given paragraph."""
        new_p = OxmlElement("w:p")
        paragraph._p.addnext(new_p)
        new_para = Paragraph(new_p, paragraph._parent)
        if text:
            new_para.add_run(text)
        if style is not None:
            new_para.style = style
        return new_para
    except:
        print("Error when inserting paragraph")


def delete_paragraph(paragraph):
    p = paragraph._element
    p.getparent().remove(p)
    paragraph._p = paragraph._element = None


def in_which_run_is(m, p):
    # Check if 'm' is outside the bounds of the paragraph's text.
    if m < 0 or m >= len(p.text):
        return None

    cumulative_length = 0
    # Enumerate through runs to get both index and the run itself.
    for index, run in enumerate(p.runs):
        cumulative_length += len(run.text)
        # If the cumulative length after adding this run includes 'm', return index.
        if cumulative_length > m:
            return index

    # If 'm' is not found in any runs, though this should be caught by the initial check.
    return None


def at_which_position_in_its_run_is(m, p):
    # Validate 'm' to ensure it is within the bounds of the paragraph.
    if m < 0 or m >= len(p.text):
        return None

    cumulative_length = 0  # Holds the cumulative length of text up to the current run.
    for _, run in enumerate(p.runs):
        previous_cumulative_length = cumulative_length
        cumulative_length += len(run.text)

        # If the cumulative length now includes 'm', calculate its position in this run.
        if cumulative_length > m:
            # 'm' minus the cumulative length of all previous runs gives the position in the current run.
            return m - previous_cumulative_length

    # In case 'm' is somehow outside the total text length (which should be caught by the initial check).
    return None


def cp(m, n, p_src, p_dest):
    if m < 0 or n > len(p_src.text):
        return None
    r_start = in_which_run_is(m, p_src)
    r_finish = in_which_run_is(n, p_src)

    if r_start == None or r_finish == None:
        return None

    for i in range(r_start, r_finish + 1):
        run = p_src.runs[i]
        r_copy = copy.deepcopy(run)._r

        a = 0
        o = len(run.text)
        if i == r_start:
            a = at_which_position_in_its_run_is(m, p_src)
        if i == r_finish:
            o = at_which_position_in_its_run_is(n, p_src) + 1

        r_copy.text = r_copy.text[a:o]
        p_dest._p.append(r_copy)


def remove_run(run, p):
    i = len(p.runs) - 1
    while i >= 0:
        if p.runs[i]._r == run._r:
            p._p.remove(p.runs[i]._r)
            return run
        i -= 1
    return None


def rm(m, n, p):
    if m < 0 or m > len(p.text) - 1 or n < 0 or n > len(p.text) - 1:
        return

    r_start = in_which_run_is(m, p)
    r_finish = in_which_run_is(n, p)

    if r_start == None or r_finish == None:
        return None

    a = -1
    o = -1
    arr = []

    for i in range(r_start, r_finish + 1):
        run = p.runs[i]

        if i == r_start:
            a = at_which_position_in_its_run_is(m, p)
        if i == r_finish:
            o = at_which_position_in_its_run_is(n, p)
        if i > r_start and i < r_finish:
            arr.append(run)

    if r_start == r_finish:
        p.runs[r_start].text = p.runs[r_start].text[:a] + p.runs[r_finish].text[o + 1 :]
    else:
        p.runs[r_start].text = p.runs[r_start].text[:a]
        p.runs[r_finish].text = p.runs[r_finish].text[o + 1 :]
    for run in reversed(arr):
        remove_run(run, p)

    return p


def mv(m, n, p_src, p_dest):
    cp(m, n, p_src, p_dest)
    rm(m, n, p_src)


def ins_into_paragraph(str, m, p):
    l = 0

    if len(p.runs) == 0:
        p.text = p.text[:m] + str + p.text[m:]
        return p

    for r in p.runs:
        l += len(r.text)
        if m <= l:
            r.text = r.text[:m] + str + r.text[m:]
            break
    return p


def replace_text_in_doc(m, p_start, n, p_end, doc, text):
    """
    Ersetzt einen Textabschnitt eines Docx-Dokument durch einen übergebenen String.

    :param m: Index des Beginns des Textabschnitts im Absatz mit dem Index p_start.
    :param p_start: Index des Absatzen in dem der zu ersetzende Textabschnitt beginnt.
    :param n: Index des Endes des Textabschnitts im Absatz mit dem Index p_end.
    :param p_end: Index des Absatzen in dem der zu ersetzende Textabschnitt endet.
    :param doc: Docx-Dokument in dem der Textabschnitt ersetzt werden soll.
    :param text: Text als String der in das Dokument eingesetzt werden soll.
    """
    if p_start == p_end:
        rm(m, n, doc.paragraphs[p_start])
    else:
        p = doc.paragraphs[p_start]
        rm(m, len(p.text) - 1, p)

        for _ in range(p_start + 1, p_end):
            p = doc.paragraphs[p_start + 1]
            delete_paragraph(p)

        p = doc.paragraphs[p_start + 1]
        rm(0, n - 1, p)

    ins_into_paragraph(text, m, doc.paragraphs[p_start])
