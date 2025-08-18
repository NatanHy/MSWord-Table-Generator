import pandas as pd
from docx import Document
import docx.document
from typing import Iterable, Dict, Tuple
from config import DSL_FILE_PATH
from table_generation.table_generator import generate_table_in_document
from word_sync.heading_tree import build_heading_tree
from table_generation.table import TableCollection
import time, queue, sys, threading
from utils.redirect_manager import redirect_stdout_to
from utils.xls_parsing import *
from utils.formatting import copy_document_styles
from utils.xml import remove_table_after_heading, delete_paragraph
from docx.text.paragraph import Paragraph

class _ComponentElement:
    """
    Wrapper class to encapsulate a Component and the oxml element where this component should
    be placed in a word file.
    """
    def __init__(self, component : Component, paragraph : Paragraph):
        self.component = component
        self.paragraph = paragraph

class AsyncTableGenerator:
    """
    Class for generating tables asynchronously. Generated tables are placed in a queue provided during initiation. 
    """
    def __init__(self, queue : queue.Queue, stdout_redirect=None, template_file_path=None):
        """
        ### Parameters
        queue : `Queue` where generated tables will be placed.\n
        stdout_redirect : optional redirect for stdout \n
        template_file_path : optional file to use as a template
        """
        self.thread = None
        self.queue = queue
        self.template_file_path = template_file_path

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

        self.thread = threading.Thread(target=task)
        self.thread.start()

    def generate_tables(self, xls_paths: Iterable[str]):
        """
        Start a thread for generating tables. 
        """
        def task():
            with open(DSL_FILE_PATH, "r") as f:
                self._code = f.read()

            # Using context manager to redirect stdout
            with redirect_stdout_to(self.stdout_redirect):
                for xls_path in xls_paths:
                    self._process_file(xls_path)

        self.thread = threading.Thread(target=task)
        self.thread.start()

    def _process_file(self, xls_path: str):
        print(f"Parsing {xls_path}...")

        xls = pd.ExcelFile(xls_path)

        components = parse_components(xls)
        variable_names = parse_variables(xls)

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
        # Headings under which the tables should be generated
        headings = root.filter(lambda node: node.heading is not None and node.heading.text == "Dependencies between processes and variables")
        
        parsed_components = {}
        variables = {}

        filtered_components = []
        for heading in headings:
            # process type used to find correct excel file
            process_type = heading.get_parent_heading_absolute(1).text #type: ignore
            # component name used to find correct component in excel file
            component_name = heading.get_parent_heading_relative(1).text #type: ignore

            # Ignore process type if it is not defined in the document
            if (xls_path := get_xls_from_process_type(process_type, xls_paths)) is None:
                continue

            # Use cahced components
            if xls_path in parsed_components:
                components = parsed_components[xls_path]
            else:
                # Parse components and variables for unseen file
                xls_file = parse_excel_cached(xls_path)

                components = parse_components(xls_file)
                for k, v in parse_variables(xls_file).items():
                    variables[k] = v
            
            for c in components:
                if c.name == component_name:
                    para = heading.get_or_insert_paragraph(-1)
                    component_element = _ComponentElement(c, para) #type: ignore
                    filtered_components.append(component_element)
                    break
        
        return filtered_components, variables

        
