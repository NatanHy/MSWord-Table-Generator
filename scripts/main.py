import pandas as pd
from geosphere import GeoSphere
from typing import List, Dict
from table_generator import generate_table
import time
import sys
from os import getcwd

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
    print("Parsing Excel file...")

    try:
        xls_path = sys.argv[1]
    except ValueError:
        raise ValueError("No file provided")

    if not xls_path.endswith(".xlxs") or xls_path.endswith(".xls"):
        raise ValueError("Unexpected file type. Provided file must be an excel file.")
    
    try:
        output_dir = sys.argv[2]
    except:
        print(f"No output directory specified, defaulting to {getcwd()}/files")
        output_dir = "files"

    try:
        xls = pd.ExcelFile(xls_path)
    except FileNotFoundError as e:
        raise e

    geospheres = parse_geospheres(xls)
    variable_descriptions = parse_variables(xls)

    print("Done.")
    print("Generating Word tables...")

    geospheres = geospheres[:1]
    successful = 0
    for i, geosphere in enumerate(geospheres):
        try:
            start = time.time()
            generate_table(geospheres[0], variable_descriptions, f"{output_dir}/table_{geosphere.id}.docx")
            end = time.time()
            successful += 1
            print(f"    Generated table for {geosphere.id} : {successful} / {len(geospheres)} | {end - start:.2f}s")
        except Exception as e:
            print(f"    Failed to generate table for {geosphere.id} : {e}")
    print(f"Operation completed. Generated {successful} files in {output_dir}/")