from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

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