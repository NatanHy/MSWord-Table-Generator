from table_generation.geosphere import GeoSphere, GeoSphereInfo
from typing import Tuple, Dict, List
from table_generation.fixed_table import FixedTable
from docx import Document
from docx.table import _Cell
import docx.document
from table_generation.table_config import configure_document, configure_table
from table_generation.parser import Parser
from utils.clean_strings import format_raw_value
import time

def span_text_over_cells(table : FixedTable, pos1 : Tuple[int, int], pos2 : Tuple[int, int], text : str):
    cell = table.cell(*pos1).merge(table.cell(*pos2))
    cell.text = text

def set_text(table : FixedTable, pos : Tuple[int, int], text : str):
    cell = table.cell(*pos)
    cell.text = text

def add_info(table : FixedTable, info : GeoSphereInfo, variable_descriptions : Dict[str, str]):
    variables = info.indicies(0)
    num_periods = info.num_time_periods()

    time_periods = info.indicies(2)
    row = 2

    for var in variables:
        desc = variable_descriptions[var]
        for i in range(num_periods):
            time_period = time_periods[i + 1]

            table.cell(row, 0).text = desc    
            table.cell(row, 1).text = \
                info.get_value(var, "Variable influence on process", "Influence present?", "Yes/No") + \
                "\n" + \
                info.get_value(var, "Variable influence on process", "Influence present?", "Description")
            table.cell(row, 2).text = time_period
            table.cell(row, 3).text = \
                info.get_value(var, "Variable influence on process", time_period, "Rationale")
            
            table.cell(row, 4).text = \
                info.get_value(var, "Process influence on variable", "Influence present?", "Yes/No") + \
                "\n" + \
                info.get_value(var, "Process influence on variable", "Influence present?", "Description")
            table.cell(row, 5).text = time_period
            table.cell(row, 6).text = \
                info.get_value(var, "Process influence on variable", time_period, "Rationale")
            
            row += 1

def _get_col_sequences(table : FixedTable, col : int, force_cutoffs) -> List[Tuple[_Cell, _Cell]]:
    start = 0
    cur = 1
    start_cell = table.cell(start, col)
    seqs = []

    # Iterate over rows
    while cur < table.rows:
        cell = table.cell(cur, col)

        # Found a row with different text, or a forced cutoff
        if cell.text != start_cell.text or cur in force_cutoffs:
            # If more than 1 consecutive cell with identical text, merge
            if cur - start > 1:
                end_cell = table.cell(cur - 1, col)
                seqs.append((start_cell, end_cell))
            start = cur # Next sequence starts at cur
            start_cell = cell
        cur += 1

    # Run one more check to include last sequence
    if cur - start > 1:
        end_cell = table.cell(cur - 1, col)
        seqs.append((start_cell, end_cell))

    return seqs

def merge_table_rows(table : FixedTable, force_cutoffs=[]):
    for col in range(table.cols):
        seqs = _get_col_sequences(table, col, force_cutoffs)
        for start_cell, end_cell in seqs:
            text_before_merge = start_cell.text
            start_cell.merge(end_cell)
            start_cell.text = text_before_merge

def generate_document(geosphere : GeoSphere, variable_descriptions : Dict[str, str], code : str, parser=Parser()) -> docx.document.Document:
    """
    Generates a word document with a table specifying information for the given geosphere. 
    """

    start_tot = time.time()

    word_document = Document()
    configure_document(word_document) # User specified document configuration
    info = geosphere.get_info()

    # Parse config file

    parser.parse(code)
    table_state = parser.execute(info, variable_descriptions)

    table = FixedTable(word_document, table_state.rows, table_state.cols)

    for i in range(table_state.rows):
        for j in range(table_state.cols):
            table.cell(i, j).text = format_raw_value(table_state.arr[i][j])

    for span in table_state.spans:
        cell1 = table.cell(*span.pos1)
        cell2 = table.cell(*span.pos2)
        cell1.merge(cell2)
        cell1.text = span.text

    end = time.time()

    start = time.time()
    merge_table_rows(table, force_cutoffs=table_state.force_cutoffs)
    end = time.time()
    print(f"merge: {end - start:.3f}")

    configure_table(table) # User specified table configuration
    
    end_tot = time.time()
    print(f"Total: {end_tot - start_tot:.3f}")
    return word_document