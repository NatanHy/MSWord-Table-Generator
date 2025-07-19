from typing import Any, Dict
from docx.table import _Cell
from docx.shared import Cm, Pt
import docx.document
from table_generation.fixed_table import FixedTable
from config.document_config import * # Constants
import ast

def format_document(doc : docx.document.Document): 
    sections = doc.sections

    for section in sections:
        section.left_margin = Cm(LEFT_MARGIN)
        section.right_margin = Cm(RIGHT_MARGIN)

def format_table(table : FixedTable):
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
        
def style(cell : _Cell, style : str):
    paragraphs = cell.paragraphs
    for paragraph in paragraphs:
        for run in paragraph.runs:
            _apply_font_attributes(run.font, style)

def _apply_font_attributes(font, attr_string):
    pairs = attr_string.split(',')
    for pair in pairs:
        if not pair.strip():
            continue
        try:
            key, value_expr = pair.split('=', 1)
            key = key.strip()

            # Raw eval, if we ever expect an untrusted user to modify table.dsl, this should be changed
            value = eval(value_expr.strip())
            setattr(font, key, value)
        except Exception as e:
            print(f"Error processing '{pair}': {e}")

