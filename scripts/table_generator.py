from geosphere import GeoSphere, GeoSphereInfo
from typing import List, Tuple, Dict
from docx import Document
from docx.table import Table
from table_config import *
from table_config import configure_document, configure_table



def span_text_over_cells(table : Table, pos1 : Tuple[int, int], pos2 : Tuple[int, int], text : str):
    cell = table.cell(*pos1).merge(table.cell(*pos2))
    cell.text = text

def set_text(table : Table, pos : Tuple[int, int], text : str):
    cell = table.cell(*pos)
    cell.text = text

def make_bold(table : Table, pos1 : Tuple[int, int], pos2 : Tuple[int, int]):
    for i in range(pos1[0], pos2[0]):
        for j in range(pos1[1], pos2[1]):
            cell = table.cell(i, j)
            paragraphs = cell.paragraphs
            for paragraph in paragraphs:
                for run in paragraph.runs:
                    run.font.bold = True
            

def add_info(table : Table, geosphere : GeoSphere, variable_descriptions : Dict[str, str]):
    info = geosphere.get_info()

    variables = info.indicies(0)
    num_periods = info.num_time_periods()

    time_periods = info.indicies(2)

    for var in variables:
        desc = variable_descriptions[var]
        for i in range(num_periods):
            table.add_row()
            time_period = time_periods[i + 1]

            table.cell(-1, 0).text = desc    
            table.cell(-1, 1).text = \
                info.get_value(var, "Variable influence on process", "Influence present?", "Yes/No") + \
                "\n" + \
                info.get_value(var, "Variable influence on process", "Influence present?", "Description")
            table.cell(-1, 2).text = time_period
            table.cell(-1, 3).text = \
                info.get_value(var, "Variable influence on process", time_period, "Rationale")
            
            table.cell(-1, 4).text = \
                info.get_value(var, "Process influence on variable", "Influence present?", "Yes/No") + \
                "\n" + \
                info.get_value(var, "Process influence on variable", "Influence present?", "Description")
            table.cell(-1, 5).text = time_period
            table.cell(-1, 6).text = \
                info.get_value(var, "Process influence on variable", time_period, "Rationale")

def merge_table_rows(table : Table):
    for col in table.columns:
        prev_cell = None
        start_cell_indx = 0

        for i, cell in enumerate(col.cells):
            if prev_cell is not None and cell.text == prev_cell.text:
                if i == len(col.cells) - 1:
                    if i != start_cell_indx:
                        text_before_merge = prev_cell.text
                        col.cells[i].merge(col.cells[start_cell_indx])
                        prev_cell.text = text_before_merge
                continue
            else:
                if prev_cell is None:
                    start_cell_indx = i
                    prev_cell = cell
                    continue

                if i != start_cell_indx + 1:
                    text_before_merge = prev_cell.text
                    col.cells[i - 1].merge(col.cells[start_cell_indx])
                    prev_cell.text = text_before_merge

                start_cell_indx = i
                prev_cell = cell

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

    merge_table_rows(table)
    configure_table(table)
    make_bold(table, (0, 0), (2, 7))

    word_document.save("files/word/test.docx")