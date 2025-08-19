import pandas as pd
from dataclasses import dataclass
from typing import List
from utils.formatting import format_raw_value
from utils.dataframes import *
from functools import cache

VAR_COL = "C" # Column where variables are e.g. VarGe01
DESC_ROW = 18 # Row of Yes/No, Description, How, Rationale
VAR_ROW = DESC_ROW + 1  # Row of top-most variable (VarGe01)
VIP_ROW = 17  # Row of time periods for "Variable influence on process"

# Top left cell of the input area (where the data is located)
VAR_INF_COL = "F"
VAR_INF_ROW = VAR_ROW

def var_to_offset(var : str) -> int:
    number = int(var[-2:])
    return number - 1

@dataclass
class Component:
    xls : pd.ExcelFile
    id : str
    name : str
    system_component : str

    def get_info(self) -> 'ComponentInfo':
        return ComponentInfo(self.id, self.xls)

class ComponentInfo:
    def __init__(self, id : str, xls : pd.ExcelFile):
        from utils.xls_parsing import get_filtered_by_id
        self.df = xls.parse(f"{id}_INF", header=None) # Drop headers since excel file is not structured like a dataframe
        self.variables = get_filtered_by_id(xls, "Var").iloc[:, 0].values.tolist()
    
    @property
    def influences(self) -> List[str]:
        return self.indicies(1)
    
    @property
    def time_periods(self) -> List[str]:
        return self.indicies(2)[1:]

    def num_time_periods(self) -> int:
        # -1 to not include "Influence present?" header as a time period
        return len(self.time_periods)

    def num_variables(self) -> int:
        return len(self.variables)

    def indicies(self, level : int) -> List[str]:
        """
        Returns the values can be used as indicies for a specified level.
        """
        match level:
            case 0:
                return self.variables
            case 1:
                return ["Variable influence on process", "Process influence on variable"]
            case 2:
                return get_non_null_values_from_row(self.df, VIP_ROW).to_list()
            case 3:
                values = get_non_null_values_from_row(self.df, DESC_ROW).to_list()[1:]
                return list(set(values))
            case _:
                raise ValueError(f"Level {level} is not a valid index.")
    
    @cache
    def _get_l0_df(self, l0) -> pd.DataFrame:
        n = var_to_offset(l0)
        # i, j is the index of the "top-left" item for the given variable
        j, i = excel_to_indx(VAR_INF_COL, VAR_INF_ROW)

        # Row offsets from "Variable influence on process" to "Process influence on variable"
        piv_offset = self.num_variables() + 4

        num_time_periods = self.num_time_periods() 

        row_indices = [i - 2, i - 1, i + n, i + n + piv_offset]
        col_range = set(range(j, j + 3 * num_time_periods + 2))
        col_exclude = set([j + 2 + 3*i for i in range(num_time_periods)])

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
            return l1_df.iloc[1:, l1_df.columns.get_indexer(cols_to_keep)] #type: ignore
        except IndexError:
            raise ValueError(f"\nInvalid level 2 index {l2}, valid values are {self.indicies(2)}")

    @cache
    def _get_l3_df(self, l0, l1, l2, l3) -> pd.DataFrame:
        l2_df = self._get_l2_df(l0, l1, l2)
        if l3 not in self.indicies(3):
            raise ValueError(f"Invalid level 3 index {l3}, valid values are {self.indicies(3)}")

        df = make_first_row_headers(l2_df)
        return df[l3]

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
