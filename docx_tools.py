import copy
from docx.oxml import OxmlElement
from docx.text.paragraph import Paragraph


def combineDocText(doc):
    """
    Kombiniert den äußeren und inneren Text eines Dokumentes und gibt diesen zurück.

    Parameters:
    - param doc: Das Dokument, von dem der Text extrahiert werden soll.
      Dies könnte ein Objekt sein, das sowohl `doc_outer_text` als auch `doc_inner_text` Methoden/Funktionen unterstützt.

    Return:
    - return: Ein String, der die Kombination aus dem äußeren und inneren Text des Dokumentes darstellt.

    Raises:
    - raises AttributeError: Wenn das übergebene `doc`-Objekt nicht die erforderlichen Methoden `doc_outer_text` und `doc_inner_text` unterstützt.
    """
    return extractOuterDocText(doc) + extractInnerDocText(doc)


def extractOuterDocText(doc):
    """
    Extrahiert den äußeren Text eines Dokumentes, indem es den Text jedes Paragraphen sammelt und zusammenfügt.

    Parameters:
    - param doc: Das Dokument, von dem der äußere Text extrahiert werden soll.
      Dies könnte ein Objekt sein, das eine Eigenschaft `paragraphs` hat, welche eine Liste von Paragraph-Objekten ist,
      wobei jedes Paragraph-Objekt eine Eigenschaft `text` hat, die den Text des Paragraphen als String enthält.

    Return:
    - return: Ein String, der den gesamten äußeren Text des Dokumentes darstellt.

    Raises:
    - raises AttributeError: Wenn das übergebene `doc`-Objekt nicht die erforderliche Eigenschaft `paragraphs` besitzt oder
      wenn ein Paragraph-Objekt innerhalb von `doc.paragraphs` nicht die Eigenschaft `text` besitzt.
    """
    fullText = ""
    for p in doc.paragraphs:
        fullText += p.text
    return fullText


def extractInnerDocText(doc):
    """
    Extrahiert den inneren Text eines Dokumentes durch Iteration über alle Tabellen und sammelt deren Inhalte.

    Parameters:
    - param doc: Das Dokument, von dem der innere Text extrahiert werden soll.
      Das Objekt sollte eine Funktion oder Methode `iterate_tables` unterstützen, die die Iteration durch alle Tabellen im Dokument ermöglicht und deren Inhalte sammelt.

    Return:
    - return: Ein String, der den gesamten inneren Text des Dokumentes darstellt, basierend auf den Inhalten der Tabellen.

    Raises:
    - raises AttributeError: Wenn das übergebene `doc`-Objekt nicht die erforderliche Methode `iterate_tables` besitzt.
    """
    return concatTableTexts(doc)


def concatCellTexts(row, func=lambda cell: cell.text):
    """
    Iterates over cells in a given row, applying a function to each cell and concatenating the results.

    Parameters:
    - row: The row object containing cells to iterate over.
    - func: A function that takes a cell object as input and returns a string. Defaults to extracting the cell's text.

    Returns:
    - A string that is the concatenation of the function's results for each cell in the row, with a newline character after each cell's output.
    """
    full_text = ""
    for cell in row.cells:
        full_text += func(cell) + "\n"
        if len(cell.tables) > 0:
            full_text += concatTableTexts(cell, func)
    return full_text


def concatRowTexts(table, func=lambda cell: cell.text):
    """
    Iterates over rows in a given table, applying the iterate_cells function to each row.

    Parameters:
    - table: The table object containing rows to iterate over.
    - func: A function to be passed to iterate_cells, which is applied to each cell.

    Returns:
    - A string that is the concatenation of all text from all cells in all rows of the table.
    """
    full_text = ""
    for row in table.rows:
        full_text += concatCellTexts(row, func)
    return full_text


def concatTableTexts(node, func=lambda cell: cell.text):
    """
    Iterates over tables in a given node (e.g., a document or another table cell), applying the iterate_rows function to each table.

    Parameters:
    - node: The node object containing tables to iterate over.
    - func: A function to be passed to iterate_rows, which is then applied to each cell in each row of each table.

    Returns:
    - A string that is the concatenation of all text from all cells in all tables within the node.
    """
    full_text = ""
    if len(node.tables) > 0:
        for table in node.tables:
            full_text += concatRowTexts(table, func)
    return full_text


def duplicatePara(para):
    """
    Dupliziert ein gegebenes Objekt `p` tiefgehend und fügt das duplizierte Objekt in die Sequenz direkt nach `p` ein.

    Parameters:
    - param p: Das Objekt, das dupliziert werden soll.
      Dieses Objekt muss ein Attribut `_p` haben, welches wiederum die Methode `addnext` unterstützen muss, um das duplizierte Objekt in die Sequenz einfügen zu können.

    Return:
    - return: Das duplizierte Objekt `p_new`.

    Raises:
    - raises AttributeError: Wenn das übergebene Objekt `p` nicht das erforderliche Attribut `_p` oder die Methode `addnext` besitzt.
    - raises CopyError: Wenn der tiefe Kopiervorgang (deep copy) fehlschlägt.
    """

    p_new = copy.deepcopy(para)
    para._p.addnext(p_new._p)
    return p_new


def appendPara(para, txt=None, style=None):
    """
    Fügt einen neuen Absatz nach einem gegebenen Absatz hinzu. Optional kann der Text und der Stil des neuen Absatzes spezifiziert werden.

    Parameters:
    - param paragraph: Das Absatzobjekt, nach dem der neue Absatz eingefügt werden soll.
      Es wird erwartet, dass dieses Objekt ein `_p` Attribut für den aktuellen Absatz und ein `_parent` Attribut für das Elternobjekt hat.
    - param text: Optionaler Text, der dem neuen Absatz hinzugefügt wird. Default ist None.
    - param style: Optionaler Stil, der auf den neuen Absatz angewendet wird. Default ist None.

    Return:
    - return: Das Objekt des neu eingefügten Absatzes.

    Raises:
    - raises Exception: Wenn beim Einfügen des neuen Absatzes ein Fehler auftritt, wird eine allgemeine Exception geworfen.
    """
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
    """
    Bestimmt den Index des Textlaufs (`run`), in dem sich das Zeichen an Position 'm' im Text des Absatzes 'p' befindet.

    Parameters:
    - param m: Die Position des Zeichens im Gesamttext des Absatzes, für die der zugehörige Textlauf ermittelt werden soll.
    - param p: Das Absatzobjekt, das die Textläufe (`runs`) enthält. Es wird erwartet, dass dieses Objekt eine Eigenschaft `text` für den Gesamttext
      und eine Liste `runs` für die Textläufe hat.

    Return:
    - return: Der Index des Textlaufs, der das Zeichen an Position 'm' enthält, oder `None`, wenn 'm' außerhalb der Grenzen des Absatztextes liegt
      oder kein entsprechender Textlauf gefunden wird.

    Raises:
    - Es werden keine Exceptions direkt von dieser Funktion ausgelöst, aber sie gibt `None` zurück, wenn die Bedingungen nicht erfüllt sind.
    """
    # Überprüft, ob 'm' außerhalb der Grenzen des Absatztextes liegt.
    if pos < 0 or pos >= len(para.text):
        return None

    cumulative_length = 0
    # Durchläuft die Textläufe, um sowohl den Index als auch den Textlauf selbst zu erhalten.
    for index, run in enumerate(para.runs):
        cumulative_length += len(run.text)
        # Wenn die kumulative Länge nach Hinzufügen dieses Textlaufs 'm' einschließt, wird der Index zurückgegeben.
        if cumulative_length > pos:
            return index

    # Wenn 'm' in keinem der Textläufe gefunden wird, obwohl dies durch die anfängliche Überprüfung abgefangen werden sollte.
    return None


def findPosInRun(pos, para):
    if pos < 0 or pos >= len(para.text):
        return None

    cumulative_length = (
        0
    )
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
            insert_position = pos - (
                l - len(r.text)
            )
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
