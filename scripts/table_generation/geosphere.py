import pandas as pd
from dataclasses import dataclass
from openpyxl.utils import column_index_from_string
from typing import Tuple, List, Any
from utils.clean_strings import format_raw_value
from functools import cache

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

def make_first_row_headers(df) -> pd.DataFrame:
    # Promote first row to header
    df.columns = df.iloc[0]         # Set first row as header
    return df[1:]                     # Drop the first row (now header)

def var_to_offset(var : str) -> int:
    return int(var.removeprefix("VarGe")) - 1

class GeoSphereInfo:
    def __init__(self, geosphere_id : str, xls : pd.ExcelFile):
        self.df = xls.parse(f"{geosphere_id}_INF", header=None) # Drop headers since excel file is not structured like a dataframe

    @property
    def variables(self) -> List[str]:
        return self.indicies(0)
    
    @property
    def influences(self) -> List[str]:
        return self.indicies(1)
    
    @property
    def time_periods(self) -> List[str]:
        return self.indicies(2)[1:]

    def num_time_periods(self) -> int:
        # -1 to not include "Influence present?" header as a time period
        return len(self.time_periods) - 1

    def num_variables(self) -> int:
        return len(self.variables)

    def indicies(self, level : int) -> List[str]:
        """
        Returns the values can be used as indicies for a specified level.
        """
        match level:
            case 0:
                # Extract column with just variable names as a flat list
                return get_cell_range("C", 19, 1, 13, self.df).iloc[:, 0].tolist()
            case 1:
                return ["Variable influence on process", "Process influence on variable"]
            case 2:
                vip = get_non_null_values_from_row(self.df, 17).to_list()
                piv = get_non_null_values_from_row(self.df, 34).to_list()
                return vip if len(vip) > len(piv) else piv
            case 3:
                values = get_non_null_values_from_row(self.df, 18).to_list()[1:]
                return list(set(values))
            case _:
                return []
    
    @cache
    def _get_l0_df(self, l0) -> pd.DataFrame:
        n = var_to_offset(l0)
        # i, j is the index of the "top-left" item for the given variable
        j, i = excel_to_indx("F", 56)

        # Row offsets from "Variable influence on process" to "Process influence on variable"
        piv_offset = 17

        row_indices = [i - 2, i - 1, i + n, i + n + piv_offset]
        col_range = set(range(j, j + 14))
        col_exclude = set([j + 2, j + 5, j + 8, j + 11])

        # Extract rows from the DataFrame
        df = self.df.iloc[row_indices, list(col_range - col_exclude)]
        return df
    
    @cache
    def _get_l1_df(self, l0, l1) -> pd.DataFrame:
        l0_df = self._get_l0_df(l0)
        match l1:
            case "Variable influence on process":
                return l0_df.iloc[[0, 1, 2]]
            case "Process influence on variable":
                return l0_df.iloc[[0, 1, 3]]
            case _:
                raise ValueError(f"Invalid level 1 index {l1}, valid values are {self.indicies(1)}")

    @cache
    def _get_l2_df(self, l0, l1, l2) -> pd.DataFrame:
        l1_df = self._get_l1_df(l0, l1)
        try:
            headers = l1_df.iloc[0]
            match_indices = [i for i, val in enumerate(headers) if val == l2]
            
            # For each match, get the index plus the next column (if it exists)
            cols_to_keep_indices = []
            for i in match_indices:
                cols_to_keep_indices.append(i)
                if i + 1 <= len(headers):
                    cols_to_keep_indices.append(i + 1)

            # Get column names based on indices
            cols_to_keep = l1_df.columns[cols_to_keep_indices]

            # Select columns (all rows)
            return l1_df.iloc[1:, l1_df.columns.get_indexer(cols_to_keep)]
        except IndexError:
            raise ValueError(f"\nInvalid level 2 index {l2}, valid values are {self.indicies(2)}")

    @cache
    def _get_l3_df(self, l0, l1, l2, l3) -> pd.DataFrame:
        l2_df = self._get_l2_df(l0, l1, l2)
        if l3 not in self.indicies(3):
            raise ValueError(f"Invalid level 3 index {l3}, valid values are {self.indicies(3)}")

        try:
            df = make_first_row_headers(l2_df)
            return df[l3]
        except:
            raise ValueError(f"Invalid level 3 index {l3}, valid values are {self.indicies(3)}")

    def get_value(self, l0, l1, l2, l3) -> str:
        """
        Get a dataframe of values in the excel file using 4-component indexing.

        ## Examples
        ```
        get_value("VarGe01", "Variable influence on process", "Temperate", "Rationale")
        ```
        """
        df = self._get_l3_df(l0, l1, l2, l3)
        return format_raw_value(df.iat[0])

@dataclass
class GeoSphere:
    xls : pd.ExcelFile
    id : str
    name : str
    description : str

    def get_info(self) -> GeoSphereInfo:
        info = GeoSphereInfo(self.id, self.xls)

        return info
