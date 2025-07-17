from typing import Any, Dict
from docx.table import _Cell
from docx.shared import Cm, Pt
import docx.document
from table_generation.fixed_table import FixedTable
from config.document_config import * # Constants

def format_document(doc : docx.document.Document): 
    sections = doc.sections

    for section in sections:
        section.left_margin = Cm(LEFT_MARGIN)
        section.right_margin = Cm(RIGHT_MARGIN)

def format_table(table : FixedTable):
    for row in range(table.rows):
        for col in range(table.cols):
            cell = table.cell(row, col)

            paragraphs = cell.paragraphs
            for paragraph in paragraphs:
                for run in paragraph.runs:
                    run.font.name = TABLE_FONT
                    font = run.font
                    font.size = Pt(TABLE_FONT_SIZE)

    table._table.style = TABLE_STYLE

def format_raw_value(val : Any) -> str:
    if val is None:
        return ""

    # Ugly nan-check, but avoids additional conversion
    match str(val):
        case "nan":
            return "—" # Note: em-dash
        case "0":
            return ""
        case res:
            return res
        
def _bold(cell : _Cell, b : bool):
    paragraphs = cell.paragraphs
    for paragraph in paragraphs:
        for run in paragraph.runs:
            run.font.bold = b

def style(cell : _Cell, style : Dict[str, Any]):
    for k, v in style.items():
        match k:
            case "bold":
                _bold(cell, v)