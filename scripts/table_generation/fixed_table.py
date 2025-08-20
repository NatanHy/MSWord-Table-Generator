from docx.table import _Cell, Table
from docx.document import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

from utils.xml import insert_table_after

class FixedTable(Table):
    """
    Class for creating a fixed-size table. Caches cells making the `cell()` 
    function much faster than in the regular `docx.table.Table` class.
    """
    def __init__(self, document : Document, rows, cols, style=None, insert_after=None):
        if insert_after is not None:
            table = insert_table_after(rows, cols, insert_after)
        else:
            table = document.add_table(rows, cols)
        
        table.style = style
        self._tbl = table._tbl
        self._parent = table._parent
        self._element = table._element

        self.document = document
        self.num_rows = rows
        self.num_cols = cols
        self._cached_cells = self._cells

    def cell(self, row_idx, col_idx) -> _Cell:
        indx = row_idx * self.num_cols + col_idx
        return self._cached_cells[indx]
    
    @property
    def width(self):
        # Read preferred width from XML if it exists

        tblW = self._tbl.tblPr.find(qn('w:tblW')) #type: ignore
        if tblW is not None:
            # Return in inches (1 inch = 1440 twips)
            return int(tblW.get(qn('w:w'))) / 1440
        return None

    @width.setter
    def width(self, size):
        self.autofit = False
        width_twips = size.twips
        tblPr = self._tbl.tblPr

        # Remove existing <w:tblW> if any
        existing = tblPr.find(qn('w:tblW')) #type: ignore
        if existing is not None:
            tblPr.remove(existing)

        # Create new <w:tblW>
        tblW = OxmlElement('w:tblW')
        tblW.set(qn('w:type'), 'dxa')
        tblW.set(qn('w:w'), str(width_twips))
        tblPr.append(tblW)