import pandas as pd
from typing import List, Dict
from table_generation import Component
from utils.dataframes import make_first_row_headers

def get_filtered_by_id(xls : pd.ExcelFile, prefix="") -> pd.DataFrame:
    # Get main sheet
    df = xls.parse("PSAR SFK FEP list", header=None)

    # Skip to row where SKB FEP ID is located
    col_b = df.columns[1]
    offset = df[df[col_b] == "SKB FEP ID"].index[0]
    df_skipped = df[offset:]

    # Filter SKB FEP ID, FEP Name and System Component columns
    df_filtered = make_first_row_headers(df_skipped)[["SKB FEP ID", "FEP Name", "System Component"]]
    var_prefix = df_filtered["SKB FEP ID"].dropna().iloc[0] # Prefix like Ge, Bio, C etc.

    # Filter variables like Ge01 or Bio01 etc.
    filtered_by_id = df_filtered[df_filtered["SKB FEP ID"].str.match(rf"{prefix}{var_prefix}[0-9]+", na=False)]
    return filtered_by_id

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