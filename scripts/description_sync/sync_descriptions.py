import pandas as pd
import openpyxl
from .get_descriptions import get_descriptions, HeadingTree
from typing import List
from table_generation import Component
from utils.xls_parsing import parse_components_cached, get_description, set_description
from utils.files import create_backup
from rapidfuzz import process, fuzz
from docx import Document
import re
from abc import ABC, abstractmethod

class _WordDescription:
    def __init__(self, node : HeadingTree):
        self.node = node
        self.component_name, self.process_type = self._get_desc_type()

    def description_paragraph(self):
        return self.node.paragraphs[0]

    def _get_desc_type(self):
        node = self.node
        component_name = None
        process_type = None

        while node.parent is not None:
            process_type = node.heading

            node = node.parent
            if component_name is None:
                component_name = node.heading
        return component_name, process_type
    
class _FileManager(ABC):
    def __init__(self, file_path : str):
        self.file_path = file_path

    @abstractmethod
    def save(self):
        pass

    def backup_and_save(self):
        create_backup(self.file_path)
        self.save()

class _ExcelFileManager(_FileManager):
    def __init__(self, file_path : str):
        super().__init__(file_path)
        self.xls = pd.ExcelFile(file_path)
        self.wb = openpyxl.load_workbook(file_path)

    def save(self):
        self.wb.save(self.file_path)

class _WordFileManager(_FileManager):
    def __init__(self, file_path : str):
        super().__init__(file_path)
        self.doc = Document(file_path)

    def save(self):
        self.doc.save(self.file_path)

class DescriptionSyncer:
    def __init__(self):
        self._process_to_xls_file = {}
        self._xls_managers = {}
        self._word_manager = None

    def sync_descriptions(self, doc_path : str, xls_file_paths : List[str]):
        self._word_manager = _WordFileManager(doc_path)

        for desc in get_descriptions(self._word_manager.doc):
            description = _WordDescription(desc)

            xls_path = self._find_best_xls_match(description, xls_file_paths)
            xls_manager = self._parse_excel(xls_path)

            components = parse_components_cached(xls_manager.xls)
            best_match, similarity = self._find_best_component_match(description, components)

            if similarity >= 80:
                self._sync_descriptions(description, best_match, xls_manager.wb)

    def save_files(self):
        if self._word_manager is not None:
            self._word_manager.backup_and_save()
        for xls_manager in self._xls_managers.values():
            xls_manager.backup_and_save()

    def _sync_descriptions(self, description : _WordDescription, component : Component, wb : openpyxl.Workbook):
        paragraph = description.description_paragraph()

        word_description = paragraph.text
        excel_description = get_description(wb, component.id)

        if len(word_description) > len(excel_description):
            set_description(wb, component.id, word_description)
        elif len(excel_description) > len(word_description):
            paragraph.text = excel_description

    def _find_best_component_match(self, description : _WordDescription, components : List[Component]):
        target = description.component_name
        choices = [comp.name for comp in components] # Extract names of components

        best_match = process.extractOne(
            target,
            choices,
            scorer=fuzz.ratio
        )

        _, similarity, index = best_match

        return components[index], similarity

    def _parse_excel(self, xls_path : str) -> _ExcelFileManager:
        if xls_path in self._xls_managers:
            return self._xls_managers[xls_path]
        
        xls_manager = _ExcelFileManager(xls_path)
        self._xls_managers[xls_path] = xls_manager
        return xls_manager

    def _find_best_xls_match(self, description : _WordDescription, xls_file_paths) -> str:
        target = description.process_type

        # If target has already been seen, return cached value
        if target in self._process_to_xls_file:
            return self._process_to_xls_file[target]

        choices = [re.split(r"[ /\\]", xls_path)[-1] for xls_path in xls_file_paths] # Only use last part of file path

        best_match = process.extractOne(
            target,
            choices,
            scorer=fuzz.token_set_ratio
        )

        # Cache and return
        best_matching_file = xls_file_paths[best_match[-1]]
        self._process_to_xls_file[target] = best_matching_file
        return best_matching_file