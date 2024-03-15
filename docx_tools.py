import copy
from docx.oxml import OxmlElement
from docx.text.paragraph import Paragraph


def combineDocText(doc):
    return extractOuterDocText(doc) + extractInnerDocText(doc)


def extractOuterDocText(doc):
    fullText = ""
    for p in doc.paragraphs:
        fullText += p.text
    return fullText


def extractInnerDocText(doc):
    return concatTableTexts(doc)


def concatCellTexts(row, func=lambda cell: cell.text):
    full_text = ""
    for cell in row.cells:
        full_text += func(cell) + "\n"
        if len(cell.tables) > 0:
            full_text += concatTableTexts(cell, func)
    return full_text


def concatRowTexts(table, func=lambda cell: cell.text):
    full_text = ""
    for row in table.rows:
        full_text += concatCellTexts(row, func)
    return full_text


def concatTableTexts(node, func=lambda cell: cell.text):
    full_text = ""
    if len(node.tables) > 0:
        for table in node.tables:
            full_text += concatRowTexts(table, func)
    return full_text


def duplicatePara(para):
    p_new = copy.deepcopy(para)
    para._p.addnext(p_new._p)
    return p_new


def appendPara(para, txt=None, style=None):
    try:
        new_p = OxmlElement("w:p")
        para._p.addnext(new_p)
        new_para = Paragraph(new_p, para._parent)
        if txt:
            new_para.add_run(txt)
        if style is not None:
            new_para.style = style
        return new_para
    except Exception as e:
        print(f"Error when inserting paragraph: {e}")


def deletePara(para):
    """
    Entfernt einen Absatz aus seinem übergeordneten Element im Dokument. Diese Funktion ändert die internen Referenzen des Absatzobjektes,
    indem sie `_element` und `_p` auf `None` setzt, um anzuzeigen, dass der Absatz nicht länger Teil des Dokuments ist.

    Parameters:
    - param paragraph: Das Absatzobjekt, das aus dem Dokument entfernt werden soll.
      Das Objekt muss über ein `_element`-Attribut verfügen, das das XML-Element des Absatzes darstellt.

    Return:
    - return: Nichts. Die Funktion verändert den Zustand des übergebenen Absatzobjektes und entfernt dessen Element aus dem Dokumentenbaum.

    Raises:
    - raises AttributeError: Wenn das übergebene Absatzobjekt nicht die erforderlichen Attribute `_element` oder `_p` besitzt.
    - raises RemoveError: Wenn das Entfernen des Absatzes aus dem Dokumentenbaum fehlschlägt, z.B. weil das Elternelement nicht gefunden werden kann.
    """
    p = para._element
    p.getparent().remove(p)
    para._p = para._element = None


def deleteTextRun(run, para):
    """
    Entfernt einen spezifischen Textlauf (`run`) aus einem Absatz (`p`). Diese Funktion durchläuft die Textläufe des Absatzes rückwärts,
    um den zu entfernenden Textlauf zu finden und ihn dann aus dem Absatz zu entfernen.

    Parameters:
    - param run: Das Textlauf-Objekt, das aus dem Absatz entfernt werden soll. Es wird erwartet, dass dieses Objekt ein Attribut `_r` hat,
      welches das zugrundeliegende XML-Element des Textlaufs repräsentiert.
    - param p: Das Absatzobjekt, aus dem der Textlauf entfernt werden soll. Das Objekt sollte eine Liste von Textläufen in `p.runs` und
      ein Attribut `_p` haben, welches das zugrundeliegende XML-Element des Absatzes repräsentiert.

    Return:
    - return: Das Textlauf-Objekt `run`, wenn es gefunden und erfolgreich entfernt wurde. Gibt `None` zurück, wenn der Textlauf im Absatz nicht gefunden wurde.

    Raises:
    - Es werden keine Exceptions direkt von dieser Funktion ausgelöst, aber durch die Verwendung von Attributen wie `_r` und `_p` besteht eine implizite
      Abhängigkeit von der Struktur des Absatz- und Textlaufobjekts, die bei Nichteinhaltung zu Fehlern führen kann.
    """
    i = len(para.runs) - 1
    while i >= 0:
        if para.runs[i]._r == run._r:
            para._p.remove(para.runs[i]._r)
            return run
        i -= 1
    return None


def findRunIndex(pos, para):
    if pos < 0 or pos >= len(para.text):
        return None

    cumulative_length = 0
    for index, run in enumerate(para.runs):
        cumulative_length += len(run.text)
        if cumulative_length > pos:
            return index

    return None


def findPosInRun(pos, para):
    if pos < 0 or pos >= len(para.text):
        return None

    cumulative_length = 0
    for _, run in enumerate(para.runs):
        previous_cumulative_length = cumulative_length
        cumulative_length += len(run.text)

        if cumulative_length > pos:
            return pos - previous_cumulative_length

    return None


def insertStrIntoPara(para, str, pos):
    l = 0

    if len(para.runs) == 0:
        para.text = para.text[:pos] + str + para.text[pos:]
        return para

    for r in para.runs:
        l += len(r.text)
        if pos <= l:
            insert_position = pos - (l - len(r.text))
            r.text = r.text[:insert_position] + str + r.text[insert_position:]
            break

    return para


def removeTextSegment(para, start, end):
    if start < 0 or start > len(para.text) or end < 0 or end > len(para.text):
        return None

    r_start = findRunIndex(start, para)
    r_finish = findRunIndex(end, para)

    if r_start == None or r_finish == None:
        return None

    a = -1
    o = -1
    arr = []

    for i in range(r_start, r_finish + 1):
        run = para.runs[i]

        if i == r_start:
            a = findPosInRun(start, para)
        if i == r_finish:
            o = findPosInRun(end, para)
        if i > r_start and i < r_finish:
            arr.append(run)

    if r_start == r_finish:
        para.runs[r_start].text = (
            para.runs[r_start].text[:a] + para.runs[r_finish].text[o + 1 :]
        )
    else:
        para.runs[r_start].text = para.runs[r_start].text[:a]
        para.runs[r_finish].text = para.runs[r_finish].text[o + 1 :]
    for run in reversed(arr):
        deleteTextRun(run, para)

    return para


def copyTextSegment(srcPara, destPara, start, end, insPos=0):
    text_to_copy = ""
    if start < 0 or end > len(srcPara.text):
        return

    for i in range(start, end + 1):
        if i < len(srcPara.text):
            text_to_copy += srcPara.text[i]

    insertStrIntoPara(destPara, text_to_copy, insPos)


def moveTextSegment(srcPara, destPara, start, end):
    copyTextSegment(srcPara, destPara, start, end)
    removeTextSegment(srcPara, start, end)


def replaceDocTextSegment(doc, startParaIdx, endParaIdx, start, end, txt):
    if startParaIdx == endParaIdx:
        removeTextSegment(doc.paragraphs[startParaIdx], start, end)
    else:
        p = doc.paragraphs[startParaIdx]
        removeTextSegment(p, start, len(p.text) - 1)

        for _ in range(startParaIdx + 1, endParaIdx):
            p = doc.paragraphs[startParaIdx + 1]
            deletePara(p)

        p = doc.paragraphs[startParaIdx + 1]
        removeTextSegment(p, 0, end)

    insertStrIntoPara(doc.paragraphs[startParaIdx], txt, start)


def extractTextBetween(doc, startParaIdx, endParaIdx, start, end):
    if startParaIdx < 0 or startParaIdx >= len(doc.paragraphs):
        raise ValueError("p_start is out of bounds")
    if endParaIdx < 0 or endParaIdx >= len(doc.paragraphs):
        raise ValueError("p_end is out of bounds")

    if startParaIdx == endParaIdx:
        return doc.paragraphs[startParaIdx].text[start:end]

    string_parts = []
    for i in range(startParaIdx, endParaIdx + 1):
        p_text = doc.paragraphs[i].text
        if i == startParaIdx:
            string_parts.append(p_text[start:])
        elif i == endParaIdx:
            string_parts.append(p_text[:end])
        else:
            string_parts.append(p_text)

    return "".join(string_parts)
