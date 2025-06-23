from pandas import ExcelFile

class GeoSphere:
    def __init__(self, xls : ExcelFile, id : str, name : str, description : str):
        self.xls = xls 
        self.id = id 