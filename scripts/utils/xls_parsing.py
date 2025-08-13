import pandas as pd
import openpyxl
from typing import List, Dict, Iterable
from table_generation import Component
from utils.dataframes import make_first_row_headers
from utils.caching import cache_on_attr
from functools import cache

@cache
def parse_excel_cached(xls_path : str) -> pd.ExcelFile:
    return pd.ExcelFile(xls_path)

def get_description(wb : openpyxl.Workbook, component_id : str) -> str:
    ws = wb[component_id]
    return ws["C14"].value

def set_description(wb : openpyxl.Workbook, component_id : str, description : str):
    ws = wb[component_id]
    ws["C14"] = description

def set_component_name(wb : openpyxl.Workbook, component : Component, text : str):
    ws = wb["PSAR SFK FEP list"]

    # Loop through column B to find row of the component
    for i, row in enumerate(ws.iter_rows(min_row=2, min_col=2, max_col=2)):
        cell = row[0]
        if cell.value == component.id:
            # Set component name in column C of the same row
            ws[f"C{i}"] = text  
            break

def get_filtered_by_id(xls : pd.ExcelFile, prefix="") -> pd.DataFrame:
    # Get main sheet
    df = xls.parse("PSAR SFK FEP list", header=None)

    # Skip to row where SKB FEP ID is located
    col_b = df.columns[1]
    offset = df[df[col_b] == "SKB FEP ID"].index[0]
    df_skipped = df[offset:]

    # Filter SKB FEP ID, FEP Name and System Component columns
    df_filtered = make_first_row_headers(df_skipped)[["SKB FEP ID", "FEP Name", "System Component", "Description"]]
    var_prefix = df_filtered["SKB FEP ID"].dropna().iloc[0] # Prefix like Ge, Bio, C etc.

    # Filter variables like Ge01 or Bio01 etc.
    filtered_by_id = df_filtered[df_filtered["SKB FEP ID"].str.match(rf"{prefix}{var_prefix}[0-9]+", na=False)]
    return filtered_by_id

@cache_on_attr('io')
def parse_components_cached(xls_file : pd.ExcelFile) -> List[Component]:
    return parse_components(xls_file)

@cache_on_attr('io')
def parse_variables_cached(xls_file : pd.ExcelFile) -> Dict[str, str]:
    return parse_variables(xls_file)

def parse_components(xls : pd.ExcelFile) -> List[Component]:
    """
    Parse Component info from the PSAR SKF FEP list sheet of the excel file. 

    ### Parameters
    xls : `pd.ExcelFile` specifying Component data in the PSAR SKF FEP list sheet

    ### Returns
    List of `Component` objects. 
    """

    filtered_by_id = get_filtered_by_id(xls)
    components = []

    for _, row in filtered_by_id.iterrows():
        id = row["SKB FEP ID"]
        name = row["FEP Name"]
        system_component = row["System Component"]
        components.append(Component(xls, id, name, system_component))

    return components

def parse_variables(xls : pd.ExcelFile) -> Dict[str, str]:
    """
    Parse Variables from the PSAR SKF FEP list sheet of the excel file. 

    ### Parameters
    xls : `pd.ExcelFile` specifying Variable data in the PSAR SKF FEP list sheet

    ### Returns
    Dictionary from variables to Component names. 
    """

    filtered_by_id = get_filtered_by_id(xls, prefix="Var")
    variables = {}

    for _, row in filtered_by_id.iterrows():
        id = row["SKB FEP ID"]
        name = row["FEP Name"]
        variables[id] = name

    return variables

def get_xls_from_process_type(process_type : str, xls_files : Iterable[str]) -> str | None: 
    #TODO
    match process_type.strip():
        case "Fuel processes":
            pth = "C:/Users/natih/OneDrive/Documents/code/python code/FEP-MSWord-Table-Generator/test/2052141 - SFK FEP-katalog för FSAR - Fuel_v0.10.xlsx"
        case "Canister processes":
            pth = "C:/Users/natih/OneDrive/Documents/code/python code/FEP-MSWord-Table-Generator/test/2052142 - SFK FEP-katalog för FSAR - Canister_v0.4.xlsx"
        case _:
            return None
    if pth in xls_files:
        return pth