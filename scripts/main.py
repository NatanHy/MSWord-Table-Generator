import pandas as pd
from geosphere import GeoSphere
from typing import List, Dict
from table_generator import generate_table

XLS_PATH = "files/excel/geosphere.xlsx"

def parse_geospheres(xls : pd.ExcelFile) -> List[GeoSphere]:
    # Get main sheet
    fep_list = xls.parse("PSAR SFK FEP list", skiprows=5)
    df = fep_list[["SKB FEP ID", "FEP Name", "Description"]]

    # Filter rows that look like Ge01, Ge02 etc.
    filtered_by_id = df[df["SKB FEP ID"].str.match(r"Ge[0-9]+", na=False)]
    
    geospheres = []

    for _, row in filtered_by_id.iterrows():
        id = row["SKB FEP ID"]
        name = row["FEP Name"]
        description = row["Description"]
        geospheres.append(GeoSphere(xls, id, name, description))

    return geospheres

def parse_variables(xls : pd.ExcelFile) -> Dict[str, str]:
    # Get main sheet
    fep_list = xls.parse("PSAR SFK FEP list", skiprows=5)
    df = fep_list[["SKB FEP ID", "FEP Name"]]

    # Filter rows that look like VarGe01, VarGe02 etc.
    filtered_by_id = df[df["SKB FEP ID"].str.match(r"VarGe[0-9]+", na=False)]

    variables = {}

    for _, row in filtered_by_id.iterrows():
        id = row["SKB FEP ID"]
        name = row["FEP Name"]
        variables[id] = name

    return variables

if __name__ == "__main__":
    xls = pd.ExcelFile(XLS_PATH)

    geospheres = parse_geospheres(xls)
    variable_descriptions = parse_variables(xls)

    generate_table(geospheres[0], variable_descriptions)