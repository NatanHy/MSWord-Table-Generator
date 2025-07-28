from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.text.paragraph import Paragraph

def insert_paragraph_after(paragraph : Paragraph, text=None, style=None):
    """Insert a new paragraph after the given paragraph."""
    new_p = OxmlElement("w:p")
    paragraph._p.addnext(new_p)
    new_para = Paragraph(new_p, paragraph._parent)
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