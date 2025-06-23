import pandas as pd
from geosphere import GeoSphere
from typing import List

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

if __name__ == "__main__":
    xls = pd.ExcelFile(XLS_PATH)

    geospheres = parse_geospheres(xls)
    for g in geospheres:
        print(g)



# word_document = Document()
# excel_document = pd.read_excel("test.xlsx")

# column_headers = excel_document.columns

# sums = [excel_document[col].sum() for col in column_headers]
# means = [excel_document[col].mean() for col in column_headers]

# table = word_document.add_table(3, len(column_headers) * 2)

# for i, header in enumerate(column_headers):
#     col_a = table.columns[2 * i]
#     col_b = table.columns[2 * i + 1]

#     merged_header = col_a.cells[0].merge(col_b.cells[0])
#     merged_header.text = header

#     col_a.cells[1].text = "Sum"
#     col_b.cells[1].text = "Mean"

#     col_a.cells[2].text = str(sums[i])
#     col_b.cells[2].text = str(means[i])

# table.style = 'Table Grid'

# word_document.save("test.docx")