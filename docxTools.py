import copy
from docx.text.paragraph import Paragraph
from docx.oxml.shared import OxmlElement


def docText(doc):
    fullText = ""
    for para in doc.paragraphs:
        fullText += para.text
    return fullText


def duplicate(p):
    p_new = copy.deepcopy(p)
    p._p.addnext(p_new._p)
    return p_new


def deleteParagraph(paragraph):
    p = paragraph._element
    p.getparent().remove(p)
    paragraph._p = paragraph._element = None


def appendParagraph(paragraph, text=None, style=None):
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
    # Validate input ranges.
    if m < 0 or n > len(p.text) - 1 or m > n:
        return

    # Find the runs where the start and end positions are located.
    r_start = in_which_run_is(m, p)
    r_finish = in_which_run_is(n, p)
    if r_start is None or r_finish is None:
        return None

    # Calculate the precise start and end positions within their respective runs.
    a_start = at_which_position_in_its_run_is(m, p)
    a_finish = at_which_position_in_its_run_is(n, p)

    # Case when both start and end positions are within the same run.
    if r_start == r_finish:
        p.runs[r_start].text = (
            p.runs[r_start].text[:a_start] + p.runs[r_finish].text[a_finish + 1 :]
        )
        return p

    # Otherwise, adjust the text in the start and end runs accordingly.
    p.runs[r_start].text = p.runs[r_start].text[:a_start]
    p.runs[r_finish].text = p.runs[r_finish].text[a_finish + 1 :]

    # Remove all runs that are fully within the range to be removed.
    for i in range(r_start + 1, r_finish):
        remove_run(p.runs[i], p)

    # This assumes `remove_run` adjusts the indexing in `p.runs` correctly.
    return p


def mv(m, n, p_src, p_dest):
    cp(m, n, p_src, p_dest)
    rm(m, n, p_src)

def ins_into_paragraph(str, m, p):
    l = 0
    for r in p.runs:
        l += len(r.text)
        if m <= l:
            r.text = r.text[:m] + str + r.text[m:]
            break
    return p