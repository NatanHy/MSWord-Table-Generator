from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.oxml.table import CT_Tbl
from typing import cast
from docx.oxml import OxmlElement
from docx.text.paragraph import Paragraph
from docx.table import Table
from typing import Union
BlockItem = Union[Paragraph, Table]

def insert_table_after(rows : int, cols : int, insert_after) -> Table:
    # Create the table (in memory, detached)
    tbl = OxmlElement('w:tbl')

    # Add tblPr (optional, for compatibility)
    tblPr = OxmlElement('w:tblPr')
    tbl.append(tblPr)

    # Add table grid (optional)
    tblGrid = OxmlElement('w:tblGrid')
    for _ in range(cols):
        gridCol = OxmlElement('w:gridCol')
        tblGrid.append(gridCol)
    tbl.append(tblGrid)

    # Add rows
    for _ in range(rows):
        tr = OxmlElement('w:tr')
        for _ in range(cols):
            tc = OxmlElement('w:tc')
            p = OxmlElement('w:p')
            tc.append(p)
            tr.append(tc)
        tbl.append(tr)

    # Insert into document
    parent = insert_after._element.getparent()
    index = parent.index(insert_after._element)
    parent.insert(index + 1, tbl)
    tbl = cast(CT_Tbl, tbl)
    return Table(tbl, insert_after._parent)

def insert_paragraph_after(item : BlockItem, text=None, style=None):
    """
    Insert a new paragraph after the given block-level item (Paragraph or Table).
    """
    # Get the correct XML element
    element = item._element

    # Create and insert new paragraph XML
    new_p = OxmlElement("w:p")
    element.addnext(new_p)

    # Wrap in Paragraph with correct parent
    new_para = Paragraph(new_p, item._parent)

    if text:
        new_para.add_run(text)
    if style is not None:
        new_para.style = style

    return new_para

def clear_document(doc):
    body = doc.element.body

    # Save the last section definition if it exists
    sectPr = None
    for el in reversed(body):
        if el.tag == qn('w:sectPr'):
            sectPr = el
            break

    # Remove all children from the body
    for el in list(body):
        body.remove(el)

    # Re-append the final section definition
    if sectPr is not None:
        body.append(sectPr)

def insert_multilevel_table_caption(paragraph, table_text):
    """
    Adds multi-level caption to table, e.g Table 1-1 or Table 1-2, by modifying the xml.
    """
    paragraph.add_run("Table ")

    # Insert STYLEREF 1 \s field for section number
    fldChar1 = OxmlElement('w:fldChar')
    fldChar1.set(qn('w:fldCharType'), 'begin')

    instrText = OxmlElement('w:instrText')
    instrText.set(qn('xml:space'), 'preserve')
    instrText.text = 'STYLEREF 1 \\s'

    fldChar2 = OxmlElement('w:fldChar')
    fldChar2.set(qn('w:fldCharType'), 'separate')

    fldChar3 = OxmlElement('w:fldChar')
    fldChar3.set(qn('w:fldCharType'), 'end')

    r = paragraph.add_run()
    r._r.append(fldChar1)
    r._r.append(instrText)
    r._r.append(fldChar2)
    r._r.append(fldChar3)

    paragraph.add_run("-")

    # Insert SEQ Table \s 1 for per-section table numbering
    fldChar1_seq = OxmlElement('w:fldChar')
    fldChar1_seq.set(qn('w:fldCharType'), 'begin')

    instrText_seq = OxmlElement('w:instrText')
    instrText_seq.set(qn('xml:space'), 'preserve')
    instrText_seq.text = 'SEQ Table \\* ARABIC \\s 1'

    fldChar2_seq = OxmlElement('w:fldChar')
    fldChar2_seq.set(qn('w:fldCharType'), 'separate')

    fldChar3_seq = OxmlElement('w:fldChar')
    fldChar3_seq.set(qn('w:fldCharType'), 'end')

    r_seq = paragraph.add_run()
    r_seq._r.append(fldChar1_seq)
    r_seq._r.append(instrText_seq)
    r_seq._r.append(fldChar2_seq)
    r_seq._r.append(fldChar3_seq)

    paragraph.add_run(f" {table_text}")