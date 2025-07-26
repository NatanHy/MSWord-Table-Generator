import pandas as pd
import openpyxl
from .get_descriptions import get_descriptions, HeadingTree
from typing import List
from table_generation import Component
from utils.xls_parsing import parse_components_cached, get_description, set_description
from rapidfuzz import process, fuzz
from docx import Document
import re

class WordDescription:
    def __init__(self, node : HeadingTree):
        self.node = node
        self.component_name, self.process_type = self._get_desc_type()

    def description_paragraph(self):
        return self.node.paragraphs[0]

    def _get_desc_type(self):
        node = self.node
        component_name = None
        process_type = None

        while node.parent is not None:
            process_type = node.heading

            node = node.parent
            if component_name is None:
                component_name = node.heading
        return component_name, process_type

def sync_descriptions(description : WordDescription, component : Component, wb : openpyxl.Workbook):
    paragraph = description.description_paragraph()

    word_description = paragraph.text
    excel_description = get_description(wb, component.id)

    if len(word_description) > len(excel_description):
        set_description(wb, component.id, word_description)
    elif len(excel_description) > len(word_description):
        paragraph.text = excel_description

def find_best_component_match(description : WordDescription, components : List[Component]):
    target = description.component_name
    choices = [comp.name for comp in components] # Extract names of components

    best_match = process.extractOne(
        target,
        choices,
        scorer=fuzz.ratio
    )

    return components[best_match[-1]] # best_match[-1] is index of best match in choices

def parse_excel_cached(xls_path : str) -> pd.ExcelFile:
    global xls_files
    if xls_path in xls_files:
        return xls_files[xls_path]
    
    f = pd.ExcelFile(xls_path)
    xls_files[xls_path] = f
    return f

def parse_workbook_cached(xls_path : str) -> openpyxl.Workbook:
    global wbs

    if xls_path in wbs:
        return wbs[xls_path]
    
    wb = openpyxl.load_workbook(xls_path)
    wbs[xls_path] = wb
    return wb

def find_best_xls_match(description : WordDescription, xls_file_paths) -> str:
    target = description.process_type
    choices = [re.split(r"[ /\\]", xls_path)[-1] for xls_path in xls_file_paths] # Only use last part of file path

    best_match = process.extractOne(
        target,
        choices,
        scorer=fuzz.token_set_ratio
    )

    return xls_file_paths[best_match[-1]]

doc = Document("C:/Users/natih/OneDrive/Desktop/internal processes/2078675 - Fuel and canister process report, FSAR version_20250519.docx")

xls_file_paths = [
    "C:/Users/natih/OneDrive/Desktop/internal processes/2052141 - SFK FEP-katalog för FSAR - Fuel_v0.10.xlsx",
    "C:/Users/natih/OneDrive/Desktop/internal processes/2052142 - SFK FEP-katalog för FSAR - Canister_v0.4.xlsx",
]

xls_files = {}
wbs = {}

for desc in get_descriptions(doc):
    description = WordDescription(desc)

    print(f"Matching:  {description.process_type} - {description.component_name}:")

    xls_path = find_best_xls_match(description, xls_file_paths)

    # Get a pandas ExcelFile for logic and openpyxl Workbook for writing
    xls = parse_excel_cached(xls_path)
    wb = parse_workbook_cached(xls_path)

    components = parse_components_cached(xls)
    best_match = find_best_component_match(description, components)
    print(best_match.id)

    sync_descriptions(description, best_match, wb)

for i, wb in enumerate(wbs.values()):  
    wb.save(f"modified{i}.xlsx")
doc.save("modified.docx")