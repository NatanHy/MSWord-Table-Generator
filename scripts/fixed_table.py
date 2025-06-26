from docx.table import _Cell
from docx.document import Document

class FixedTable:
    """
    Class for creating a fixed-size table. Caches cells making the `cell()` 
    function much faster than in the regular `docx.table.Table` class.
    """
    def __init__(self, document : Document, rows, cols):
        self.document = document
        self.rows = rows
        self.cols = cols

        self._table = self.document.add_table(rows, cols)
        self._cells = self._table._cells

    def cell(self, row, col) -> _Cell:
        indx = row * self.cols + col
        return self._cells[indx]