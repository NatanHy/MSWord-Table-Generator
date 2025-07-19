import pandas as pd
from typing import List, Dict
from table_generation.component import Component

def _get_filtered_by_id(xls : pd.ExcelFile) -> pd.DataFrame:
    # Get main sheet
    df = xls.parse("PSAR SFK FEP list", header=None)

    # Skip to row where SKB FEP ID is located
    col_b = df.columns[1]
    offset = df[df[col_b] == "SKB FEP ID"].index[0]
    df_skipped = df[offset:]

    # Filter SKB FEP ID and FEP Name
    df_filtered = make_first_row_headers(df_skipped)[["SKB FEP ID", "FEP Name"]]
    var_prefix = df_filtered["SKB FEP ID"].dropna().iloc[0] # Prefix like Ge, Bio, C etc.

    # Filter variables like Ge01 or Bio01 etc.
    filtered_by_id = df_filtered[df_filtered["SKB FEP ID"].str.match(rf"{var_prefix}[0-9]+", na=False)]
    return filtered_by_id

def parse_components(xls : pd.ExcelFile) -> List[Component]:
    """
    Parse Component info from the PSAR SKF FEP list sheet of the excel file. 

    ### Parameters
    xls : `pd.ExcelFile` specifying Component data in the PSAR SKF FEP list sheet

    ### Returns
    List of `Component` objects. 
    """

    filtered_by_id = _get_filtered_by_id(xls)
    components = []

    for _, row in filtered_by_id.iterrows():
        id = row["SKB FEP ID"]
        name = row["FEP Name"]
        components.append(Component(xls, id, name, ""))

    return components

def parse_variables(xls : pd.ExcelFile) -> Dict[str, str]:
    """
    Parse VarGe from the PSAR SKF FEP list sheet of the excel file. 

    ### Parameters
    xls : `pd.ExcelFile` specifying Variable data in the PSAR SKF FEP list sheet

    ### Returns
    Dictionary from variables to Component names. 
    """

    filtered_by_id = _get_filtered_by_id(xls)
    variables = {}

    for _, row in filtered_by_id.iterrows():
        id = row["SKB FEP ID"]
        name = row["FEP Name"]
        variables[id] = name

    return variables

def make_first_row_headers(df) -> pd.DataFrame:
    # Promote first row to header
    df.columns = df.iloc[0]         # Set first row as header
    return df[1:]        

def excel_to_indx(col : str, row : int) -> Tuple[int, int]:
    # -1 from index since python uses 0-indexing whereas excel uses 1-indexing
    return (column_index_from_string(col) - 1, row - 1)

def get_cell(excel_col : str, excel_row : int, df : pd.DataFrame) -> Any:
    return get_cell_range(excel_col, excel_row, 1, 1, df)

def get_cell_range(excel_col : str, excel_row : int, col_span : int, row_span : int, df : pd.DataFrame) -> Any:
    iloc_col, iloc_row = excel_to_indx(excel_col, excel_row)
    return df.iloc[iloc_row:iloc_row+row_span, iloc_col:iloc_col+col_span]

def get_non_null_values_from_row(df, excel_row : int) -> pd.Series:
    row_indx = excel_row - 1

    if row_indx not in df.index:
        raise ValueError("Row index not found in the DataFrame.")
    return df.loc[row_indx].dropna()