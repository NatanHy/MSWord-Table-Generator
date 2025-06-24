import pandas as pd
from dataclasses import dataclass
from openpyxl.utils import column_index_from_string
from typing import Tuple

def excel_to_iloc_indx(col : str, row : int) -> Tuple[int, int]:
    # -1 from column index since iloc uses 0-indexing whereas excel uses 1-indexing
    # -2 from row because 0-indexing and iloc skips headers
    return (column_index_from_string(col) - 1, row - 2)

def make_first_row_headers(df : pd.DataFrame):
    df_with_headers = df.iloc[1:].copy()          # skip current header row
    df_with_headers.columns = df.iloc[0].values   # set headers using first row
    return df_with_headers

def get_info_table(df : pd.DataFrame, variables : pd.DataFrame, col : str, row : int,) -> pd.DataFrame:
    start_col_idx, start_row_idx = excel_to_iloc_indx(col, row)

    end_row_idx = start_row_idx + 13
    end_col_idx = start_col_idx + 1

    # +1 to end indicies since slices are exclusive
    info_df = df.iloc[start_row_idx:end_row_idx + 1, start_col_idx:end_col_idx + 1].copy()
    concat_df = pd.concat([variables.reset_index(drop=True), info_df.reset_index(drop=True)], axis=1)

    df_with_headers = make_first_row_headers(concat_df)

    return df_with_headers

class GeoSphereInfo:
    def __init__(self, geosphere_id : str, xls : pd.ExcelFile):
        self.geosphere_id = geosphere_id

        df = xls.parse(f"{self.geosphere_id}_INF")

        # Extract column with just variable names

        var_col, var_row = excel_to_iloc_indx("C", 55)
        variables = pd.DataFrame(df.iloc[var_row:var_row + 14, var_col])

        tables = {
            "variable influence on process" : {
                "influence present" : get_info_table(df,  variables, "F", 55),
                "excavation/operation" : get_info_table(df,  variables, "I", 55),
                "temperate" : get_info_table(df,  variables, "L", 55),
                "preglacial" : get_info_table(df,  variables, "O", 55),
                "glacial" : get_info_table(df,  variables, "R", 55)
            },
            "process influence on variable" : {
                "influence present" : get_info_table(df,  variables, "F", 72),
                "excavation/operation" : get_info_table(df,  variables, "I", 72),
                "temperate" : get_info_table(df,  variables, "L", 72),
                "preglacial" : get_info_table(df,  variables, "O", 72),
                "glacial" : get_info_table(df,  variables, "R", 72)
            }
        }

        self.tables = tables

@dataclass
class GeoSphere:
    xls : pd.ExcelFile
    id : str
    name : str
    description : str

    def get_info(self) -> GeoSphereInfo:
        info = GeoSphereInfo(self.id, self.xls)

        return info
