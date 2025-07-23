import pandas as pd
from typing import Iterable
from config import DSL_FILE_PATH
from table_generation.table_generator import generate_table_in_document
from table_generation.table import TableCollection
import time, queue, sys, threading
from utils.redirect_manager import redirect_stdout_to
from utils.xls_parsing import parse_components, parse_variables
from utils.formatting import copy_document_styles

class AsyncTableGenerator:
    """
    Class for generating tables asynchronously. Generated tables are placed in a queue provided during initiation. 
    """
    def __init__(self, queue : queue.Queue, stdout_redirect=None):
        """
        ### Parameters
        queue : `Queue` where generated tables will be placed.\n
        stdout_redirect : optional redirect for stdout
        """
        self.thread = None
        self.queue = queue

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
        print(f"Parsing {xls_path}")

        xls = pd.ExcelFile(xls_path)

        components = parse_components(xls)
        variable_names = parse_variables(xls)

        print("Done.")
        print("Generating Word tables...")

        # Keep track of successfully/unsuccessfully generated tables
        successful = 0
        unsuccessful = 0

        word_document = copy_document_styles("C:/Users/natih/OneDrive/Desktop/internal processes/2078675 - Fuel and canister process report, FSAR version_20250519.docx")

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

    def _generate_table(self, word_document, component, variable_names) -> bool:
        # Try generating word document and add it to the queue
        try:
            start = time.time()
            generate_table_in_document(word_document, component, variable_names, self._code)
            end = time.time()
            print(f"    Generated table for {component.id} : Success | {end - start:.2f}s")
            return True
        except Exception as e:
            print(f"    Failed to generate table for {component.id} : {e}")
            return False