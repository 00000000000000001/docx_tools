from docx import Document
from docx_tools import extractInnerDocText

def test_finds_all_content():
    doc = Document()
    
    # Eine Tabelle im Dokument hinzufügen
    tabelle = doc.add_table(rows=3, cols=3)
    
    # Liste von Worten, die in die Tabelle eingefügt werden sollen
    worte = [
        ['Wort1', 'Wort2', 'Wort3'],
        ['Wort4', 'Wort5', 'Wort6'],
        ['Wort7', 'Wort8', 'Wort9']
    ]
    
    # Durch jede Zeile und Spalte der Tabelle iterieren und Worte einfügen
    for zeilen_index, zeile in enumerate(tabelle.rows):
        for spalten_index, zelle in enumerate(zeile.cells):
            zelle.text = worte[zeilen_index][spalten_index]

    assert 'Wort1' in extractInnerDocText(doc)
    assert 'Wort2' in extractInnerDocText(doc)
    assert 'Wort3' in extractInnerDocText(doc)
    assert 'Wort4' in extractInnerDocText(doc)
    assert 'Wort5' in extractInnerDocText(doc)
    assert 'Wort6' in extractInnerDocText(doc)
    assert 'Wort7' in extractInnerDocText(doc)
    assert 'Wort8' in extractInnerDocText(doc)
    assert 'Wort9' in extractInnerDocText(doc)