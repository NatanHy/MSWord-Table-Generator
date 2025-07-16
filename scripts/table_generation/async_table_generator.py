import pandas as pd
from typing import Iterable
from table_generation import CFG_FILE_PATH
from table_generation.table_generator import generate_document
from table_generation.table import Table
import time, queue, sys, threading
from utils.redirect_manager import redirect_stdout_to
from utils.xls_parsing import parse_geospheres, parse_variables

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

    def is_running(self) -> bool:
        return self.thread is not None and self.thread.is_alive()

    def generate_tables(self, xls_paths: Iterable[str]):
        """
        Start a thread for generating tables. 
        """
        def task():
            with open(CFG_FILE_PATH, "r") as f:
                self._code = f.read()

            # Using context manager to redirect stdout
            with redirect_stdout_to(self.stdout_redirect):
                for xls_path in xls_paths:
                    self._process_file(xls_path)

        self.thread = threading.Thread(target=task)
        self.thread.start()

    def _process_file(self, xls_path: str):
        print(f"Parsing {xls_path}")
        try:
            xls = pd.ExcelFile(xls_path)
        except FileNotFoundError as e:
            raise e

        geospheres = parse_geospheres(xls)
        variable_descriptions = parse_variables(xls)

        print("Done.")
        print("Generating Word tables...")

        # Keep track of successfully/unsuccessfully generated tables
        successful = 0
        unsuccessful = 0

        for geosphere in geospheres:
            # Abort generation if stop flag is set
            if self.stop_event.is_set():
                print("Operation terminated.")
                return

            success = self._generate_table(geosphere, variable_descriptions, xls_path)
            if success:
                successful += 1
            else:
                unsuccessful += 1

        print(f"Operation completed. Generated {successful} table(s). Success {successful} | Fail {unsuccessful}")

    def _generate_table(self, geosphere, variable_descriptions, xls_path) -> bool:
        # Try generating word document and add it to the queue
        try:
            start = time.time()
            document = generate_document(geosphere, variable_descriptions, self._code)
            self.queue.put(Table(document, xls_path, geosphere.id))
            end = time.time()
            print(f"    Generated table for {geosphere.id} : Success | {end - start:.2f}s")
            return True
        except Exception as e:
            print(f"    Failed to generate table for {geosphere.id} : {e}")
            return False