from openpyxl.utils import column_index_from_string
import pandas as pd
from typing import Tuple, Any

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