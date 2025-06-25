import pandas as pd
from dataclasses import dataclass
from openpyxl.utils import column_index_from_string
from typing import Tuple, List, Any

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

def var_to_offset(var : str) -> int:
    return int(var.removeprefix("VarGe")) - 1

class GeoSphereInfo:
    def __init__(self, geosphere_id : str, xls : pd.ExcelFile):
        self.geosphere_id = geosphere_id
        self.df = xls.parse(f"{self.geosphere_id}_INF", header=None) # Drop headers since excel file is not structured like a dataframe

    @property
    def variables(self) -> List[str]:
        # Extract column with just variable names as a flat list
        return get_cell_range("C", 19, 1, 13, self.df).iloc[:, 0].tolist()


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
                vip = get_non_null_values_from_row(self.df, 17).to_list()
                piv = get_non_null_values_from_row(self.df, 34).to_list()
                return vip if len(vip) > len(piv) else piv
            case 3:
                values = get_non_null_values_from_row(self.df, 18).to_list()[1:]
                return list(set(values))
            case _:
                return []
                
    def get_value(self, l0, l1, l2, l3):
        try:
            l0_row_offset = var_to_offset(l0)
        except Exception as e:
            raise ValueError(f"Invalid level 0 index {l0}, valid values are {self.indicies(0)}")
        
        match l1:
            case "Variable influence on process":
                # l1_row_offset = 18
                l1_row_offset = 55
            case "Process influence on variable":
                # l1_row_offset = 35
                l1_row_offset = 72
            case _:
                raise ValueError(f"Invalid level 1 index {l1}, valid values are {self.indicies(1)}")
        
        try:
            row = self.df.loc[l1_row_offset - 2]
            l2_col_offset = row[row == l2].index.to_list()[0]
        except IndexError:
            raise ValueError(f"Invalid level 2 index {l2}, valid values are {self.indicies(2)}")

        if l3 not in self.indicies(3):
            raise ValueError(f"Invalid level 2 index {l3}, valid values are {self.indicies(3)}")

        match l3:
            case "Yes/No" | "How":
                l3_col_offset = 0
            case "Description" | "Rationale":
                l3_col_offset = 1
            case _:
                raise ValueError(f"Invalid level 2 index {l3}, valid values are {self.indicies(3)}")
        
        return self.df.iloc[l0_row_offset + l1_row_offset, l2_col_offset + l3_col_offset]

@dataclass
class GeoSphere:
    xls : pd.ExcelFile
    id : str
    name : str
    description : str

    def get_info(self) -> GeoSphereInfo:
        info = GeoSphereInfo(self.id, self.xls)

        return info
