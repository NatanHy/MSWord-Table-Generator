import threading
import pandas as pd
from geosphere import GeoSphere
from typing import Dict, Iterable, List
from table_generator import generate_document
from table import Table
import time, queue, sys, io
from contextlib import contextmanager

@contextmanager
def redirect_stdout_to(redirector):
    """
    Context manager for redirectring stdout. Redirects stdout within the context, and 
    redirect it back to the previous state after. 
    
    ## Example

    ```
    with redirect_stdout_to(my_redirect_widget):
        print("Printing to my widget :)")
    ```
    """
    original = sys.stdout
    sys.stdout = redirector
    try:
        yield
    finally:
        sys.stdout = original


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

class AsyncTableGenerator:
    def __init__(self, queue : queue.Queue, stdout_redirect=None):
        self.thread = None
        self.queue = queue

        if stdout_redirect is None:
            self.stdout_redirect = sys.stdout
        else:
            self.stdout_redirect = stdout_redirect
            
        self.stop_event = threading.Event()

    def is_running(self) -> bool:
        return self.thread is not None and self.thread.is_alive()

    def generate_tables(self, xls_paths: Iterable[str]):
        def task():
            with redirect_stdout_to(self.stdout_redirect):
                for xls_path in xls_paths:
                    self._process_file(xls_path)

        self.thread = threading.Thread(target=task)
        self.thread.start()

    def _process_file(self, xls_path: str):
        try:
            xls = pd.ExcelFile(xls_path)
        except FileNotFoundError as e:
            raise e

        geospheres = parse_geospheres(xls)
        variable_descriptions = parse_variables(xls)

        print("Done.")
        print("Generating Word tables...")

        successful = 0
        unsuccessful = 0

        for geosphere in geospheres:
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
        try:
            start = time.time()
            document = generate_document(geosphere, variable_descriptions)
            self.queue.put(Table(document, xls_path, geosphere.id))
            end = time.time()
            print(f"    Generated table for {geosphere.id} : Success | {end - start:.2f}s")
            return True
        except Exception as e:
            print(f"    Failed to generate table for {geosphere.id} : {e}")
            return False