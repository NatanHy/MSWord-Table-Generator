from geosphere import GeoSphere, GeoSphereInfo
from typing import List, Tuple, Dict
from docx import Document
from docx.table import Table
from table_config import *
from table_config import configure_document



def span_text_over_cells(table : Table, pos1 : Tuple[int, int], pos2 : Tuple[int, int], text : str):
    cell = table.cell(*pos1).merge(table.cell(*pos2))
    cell.text = text

def set_text(table : Table, pos : Tuple[int, int], text : str):
    cell = table.cell(*pos)
    cell.text = text

def add_info(table : Table, geosphere : GeoSphere, variable_descriptions : Dict[str, str]):
    info = geosphere.get_info()
    for var in info.variables:
        table.add_row()
        table.cell(-1, 0).text = variable_descriptions[var] 

def generate_table(geosphere : GeoSphere, variable_descriptions : Dict[str, str]):
    word_document = Document()
    configure_document(word_document)

    table = word_document.add_table(2, 7)

    # Make header
    span_text_over_cells(table, (0, 0), (1, 0), "Variables")
    span_text_over_cells(table, (0, 1), (0, 3), "Variable influence on process")    
    span_text_over_cells(table, (0, 4), (0, 6), "Process influence on variables")  

    set_text(table, (1, 1), "Influence present? (Yes/No Description)")
    set_text(table, (1, 4), "Influence present? (Yes/No Description)")
    set_text(table, (1, 2), "Time period/Climate domain")
    set_text(table, (1, 5), "Time period/Climate domain")
    set_text(table, (1, 3), "Handling of influence \n (How/If not — Why)")
    set_text(table, (1, 6), "Handling of influence \n (How/If not — Why)")

    add_info(table, geosphere, variable_descriptions)

    table.style = "Table Grid"

    word_document.save("files/word/test.docx")