import pytest
from docx_tools import combineDocText, extractOuterDocText, extractInnerDocText, concatTableTexts, concatRowTexts, concatCellTexts

class MockCell:
    def __init__(self, text, tables=None):
        self.text = text
        self.tables = tables if tables else []

class MockRow:
    def __init__(self, cells):
        self.cells = cells

class MockTable:
    def __init__(self, rows):
        self.rows = rows

class MockDoc:
    def __init__(self, paragraphs=None, tables=None):
        self.paragraphs = paragraphs if paragraphs else []
        self.tables = tables if tables else []

@pytest.fixture
def doc_setup():
    cell1 = MockCell("Cell 1")
    cell2 = MockCell("Cell 2", tables=[MockTable([MockRow([MockCell("Nested Cell 1"), MockCell("Nested Cell 2")])])])
    row = MockRow([cell1, cell2])
    table = MockTable([row])
    paragraphs = [MockCell("Paragraph 1"), MockCell("Paragraph 2")]
    doc = MockDoc(paragraphs, [table])
    return doc

def test_extractOuterDocText(doc_setup):
    """
    Test extractOuterDocText function to ensure it correctly combines text from document paragraphs.
    """
    expected_text = "Paragraph 1Paragraph 2"
    assert extractOuterDocText(doc_setup) == expected_text

def test_extractInnerDocText(doc_setup):
    """
    Test extractInnerDocText function to ensure it correctly combines text from nested table structures in the document.
    """
    expected_text = "Cell 1\nCell 2\nNested Cell 1\nNested Cell 2\n"
    assert extractInnerDocText(doc_setup) == expected_text

def test_combineDocText(doc_setup):
    """
    Test combineDocText function to ensure it correctly combines text from both paragraphs and nested table structures.
    """
    expected_text = "Paragraph 1Paragraph 2Cell 1\nCell 2\nNested Cell 1\nNested Cell 2\n"
    assert combineDocText(doc_setup) == expected_text

def test_concatCellTexts_simple(doc_setup):
    """
    Test concatCellTexts function with a simple row to ensure it concatenates cell texts correctly.
    """
    row = doc_setup.tables[0].rows[0]  # Assuming doc_setup has at least one table with one row
    expected_text = "Cell 1\nCell 2\nNested Cell 1\nNested Cell 2\n"
    assert concatCellTexts(row) == expected_text

def test_concatRowTexts(doc_setup):
    """
    Test concatRowTexts function to ensure it concatenates text from all rows within a table correctly.
    """
    table = doc_setup.tables[0]  # Assuming doc_setup has at least one table
    expected_text = "Cell 1\nCell 2\nNested Cell 1\nNested Cell 2\n"
    assert concatRowTexts(table) == expected_text

def test_concatTableTexts(doc_setup):
    """
    Test concatTableTexts function to ensure it concatenates text from all tables within a document or node correctly.
    """
    expected_text = "Cell 1\nCell 2\nNested Cell 1\nNested Cell 2\n"
    assert concatTableTexts(doc_setup) == expected_text

def test_concatCellTexts_with_custom_function(doc_setup):
    """
    Test concatCellTexts function with a custom function to ensure it correctly applies the function to each cell's text.
    """
    row = doc_setup.tables[0].rows[0]  # Assuming doc_setup has at least one table with one row
    custom_func = lambda cell: f"Text: {cell.text}"
    expected_text = "Text: Cell 1\nText: Cell 2\nText: Nested Cell 1\nText: Nested Cell 2\n"
    assert concatCellTexts(row, custom_func) == expected_text

def test_empty_document():
    """
    Test functions with an empty document to ensure they return empty strings or behave as expected.
    """
    empty_doc = MockDoc()

    assert extractOuterDocText(empty_doc) == ""
    assert extractInnerDocText(empty_doc) == ""
    assert concatTableTexts(empty_doc) == ""

def test_document_with_no_tables():
    """
    Test functions with a document that has no tables to ensure they handle such scenarios correctly.
    """
    no_table_doc = MockDoc(paragraphs=[MockCell("Paragraph 1"), MockCell("Paragraph 2")])

    assert extractOuterDocText(no_table_doc) == "Paragraph 1Paragraph 2"
    assert extractInnerDocText(no_table_doc) == ""
    assert concatTableTexts(no_table_doc) == ""


def test_concatCellTexts_empty_row():
    """
    Test concatCellTexts function with an empty row to ensure it returns an empty string.
    """
    empty_row = MockRow([])
    assert concatCellTexts(empty_row) == ""

def test_concatRowTexts_empty_table():
    """
    Test concatRowTexts function with an empty table to ensure it returns an empty string.
    """
    empty_table = MockTable([])
    assert concatRowTexts(empty_table) == ""

def test_concatTableTexts_empty_node():
    """
    Test concatTableTexts function with a node that has no tables to ensure it returns an empty string.
    """
    empty_node = MockDoc()  # MockDoc used here for simplicity; any node-like object without tables would work
    assert concatTableTexts(empty_node) == ""

