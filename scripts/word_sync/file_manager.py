from abc import ABC, abstractmethod
from utils.files import create_backup
import pandas as pd
import openpyxl
from docx import Document

class FileManager(ABC):
    def __init__(self, file_path : str):
        self.file_path = file_path

    @abstractmethod
    def save(self):
        pass

    def backup_and_save(self):
        create_backup(self.file_path)
        self.save()

class ExcelFileManager(FileManager):
    def __init__(self, file_path : str):
        super().__init__(file_path)
        self.xls = pd.ExcelFile(file_path)
        self.wb = openpyxl.load_workbook(file_path, data_only=True)

    def save(self):
        self.wb.save(self.file_path)
        self.wb.close()

class WordFileManager(FileManager):
    def __init__(self, file_path : str):
        super().__init__(file_path)
        self.doc = Document(file_path)

    def save(self):
        self.doc.save(self.file_path)