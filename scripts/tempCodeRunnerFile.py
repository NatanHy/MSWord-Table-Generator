.ExcelFile("C:/Users/natih/Downloads/geospheres.xlsx")
geospheres = parse_geospheres(xls)
variable_descriptions = parse_variables(xls)

doc = generate_document(geospheres[0], variable_descriptions)
doc.save("test.docx")