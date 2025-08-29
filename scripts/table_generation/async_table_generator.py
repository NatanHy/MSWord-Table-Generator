import queue
import sys
import threading
import time
from typing import Iterable, Dict, Tuple, List, Callable

from docx import Document
import docx.document
from docx.text.paragraph import Paragraph
import pandas as pd

from table_generation.table_generator import generate_table_in_document
from table_generation.table import TableCollection
from table_generation.component import Component
from word_sync.heading_tree import build_heading_tree
from utils.redirect_manager import redirect_stdout_to
from utils.formatting import copy_document_styles
from utils.xml import remove_table_after_heading, delete_paragraph, parse_mappings
from utils.xls_parsing import (
    parse_components, 
    parse_variables,
    get_xls_from_component_id, 
    parse_excel_cached, 
    get_component_by_id
    )
from utils.files import ExcelFileManager, resource_path

DSL_FILE_PATH = resource_path("config/table.dsl")

class _ComponentElement:
    
    """
    Wrapper class to encapsulate a Component and the paragraph where this component should
    be placed in a word file.
    """
    def __init__(self, component : Component, paragraph : Paragraph):
        self.component = component
        self.paragraph = paragraph

class AsyncTableGenerator:
    """
    Class for generating tables asynchronously. Generated tables are placed in a queue provided during initiation. 
    """
    def __init__(self, queue : queue.Queue, stdout_redirect=None, template_file_path=None, on_fail : Callable[[Exception], None]=None):
        """
        ### Parameters
        queue : `Queue` where generated tables will be placed.\n
        stdout_redirect : optional redirect for stdout \n
        template_file_path : optional file to use as a template
        """
        self.thread = None
        self.queue = queue
        self.template_file_path = template_file_path
        self.on_fail = on_fail

        if stdout_redirect is None:
            self.stdout_redirect = sys.stdout
        else:
            self.stdout_redirect = stdout_redirect

        self.stop_event = threading.Event()
        self._code = "" # Code file will be read at runtime

    def is_done(self) -> bool:
        is_running = self.thread is not None and self.thread.is_alive()

        # If the thread is running or was stopped return false
        return (not is_running) and (not self.stop_event.is_set())

    def generate_and_insert_tables(self, xls_paths: Iterable[str], doc : docx.document.Document):
        """
        Start a thread for generating tables. 
        """
        def task():
            try:
                with open(DSL_FILE_PATH, "r") as f:
                    self._code = f.read()

                # Using context manager to redirect stdout
                with redirect_stdout_to(self.stdout_redirect):
                    print("Parsing word document...")
                    component_elements, variable_descriptions = self._parse_document(doc, xls_paths)
                    print("Done.")
                    print("Generating Word tables...")
                    for ce in component_elements:
                        if self.stop_event.is_set():
                            print("Operation terminated.")
                            return
                        # If there is already a table, remove it and it's heading
                        if remove_table_after_heading(doc, ce.paragraph.text):
                            self._generate_table(doc, ce.component, variable_descriptions, insert_after=ce.paragraph)
                            delete_paragraph(ce.paragraph)
                        else:
                            self._generate_table(doc, ce.component, variable_descriptions, insert_after=ce.paragraph)
                    print("Done.")
            except Exception as e:
                if self.on_fail:
                    self.on_fail(e)

        self.thread = threading.Thread(target=task)
        self.thread.start()

    def generate_tables(self, xls_paths: Iterable[str]):
        """
        Start a thread for generating tables. 
        """
        def task():
            try:
                with open(DSL_FILE_PATH, "r") as f:
                    self._code = f.read()

                # Using context manager to redirect stdout
                with redirect_stdout_to(self.stdout_redirect):
                    for xls_path in xls_paths:
                        self._process_file(xls_path)
            except Exception as e:
                if self.on_fail:
                    self.on_fail(e)

        self.thread = threading.Thread(target=task)
        self.thread.start()

    def _process_file(self, xls_path: str):
        print(f"Parsing {xls_path}...")

        file_manager = ExcelFileManager(xls_path)

        components = parse_components(file_manager)
        variable_names = parse_variables(file_manager)

        print("Done.")
        print("Generating Word tables...")

        # Keep track of successfully/unsuccessfully generated tables
        successful = 0
        unsuccessful = 0

        if self.template_file_path is not None:
            word_document = copy_document_styles(self.template_file_path)
        else:
            word_document = Document()

        for component in components:
            # Abort generation if stop flag is set
            if self.stop_event.is_set():
                print("Operation terminated.")
                return

            success = self._generate_table(word_document, component, variable_names)
            if success:
                successful += 1
            else:
                unsuccessful += 1

        self.queue.put(TableCollection(word_document, xls_path))
        print(f"Operation completed. Generated {successful} table(s). Success {successful} | Fail {unsuccessful}")

    def _generate_table(self, doc, component, variable_names, insert_after=None) -> bool:
        # Try generating table in the document
        try:
            start = time.time()
            generate_table_in_document(doc, component, variable_names, self._code, insert_after=insert_after)
            end = time.time()
            print(f"    Generated table for {component.id} : Success | {end - start:.2f}s")
            return True
        except Exception as e:
            print(f"    Failed to generate table for {component.id} : {e}")
            return False
        
    def _parse_document(self, doc : docx.document.Document,  xls_paths: Iterable[str]) -> Tuple[List[_ComponentElement], Dict[str, str]]:
        """
        Parse a word document for table insertion.
        """
        root = build_heading_tree(doc)
        mappings = parse_mappings(doc)
        # Headings under which the tables should be generated
        headings = root.filter(lambda node: node.heading is not None and node.heading.text == "Dependencies between processes and variables")
        
        parsed_paths = {} # Cache parsed xls files
        variables = {}    # Variable names

        filtered_components = []
        for heading in headings:
            # process type used to find correct excel file
            process_type = heading.get_parent_heading_absolute(1).text.strip() #type: ignore
            # component name used to find correct component in excel file
            component_name = heading.get_parent_heading_relative(1).text.strip() #type: ignore

            try:
                components =  mappings[process_type]
            except KeyError:
                print(f"WARNING: Missing mapping for '{process_type}', malformed mapping table?")
                continue # Trying to find a component id for non-process-type, skip iteration

            try:
                component_id = components[component_name]
            except KeyError:
                print(f"WARNING: Missing mapping for '{process_type}' - '{component_name}'.")
                continue # Trying to find a component id for non-process-type, skip iteration

            # Ignore component if it is not defined in the excel files
            if (xls_path := get_xls_from_component_id(component_id, xls_paths)) is None:
                print(f"    Could not find {component_id} in the proved excel files, skipping")
                continue
            
            # Parse variable descriptions for new xls paths
            if xls_path in parsed_paths:
                xls_file_manager = parsed_paths[xls_path]
            else:
                xls_file_manager = parse_excel_cached(xls_path)
                parsed_paths[xls_path] = xls_file_manager
                for k, v in parse_variables(xls_file_manager).items():
                    variables[k] = v # Add variable names

            component = get_component_by_id(xls_file_manager, component_id)

            para = heading.get_or_insert_paragraph(-1)
            component_element = _ComponentElement(component, para) #type: ignore
            filtered_components.append(component_element)
        return filtered_components, variables
