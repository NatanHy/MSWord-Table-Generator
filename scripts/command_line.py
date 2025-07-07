import sys
from async_table_generator import AsyncTableGenerator
import queue
from table import Table
from time import sleep

def poll_table_queue(output_dir):
    try:
        table = table_queue.get_nowait()
        table.save(output_dir)
    except queue.Empty:
        pass

if __name__ == "__main__":
    try:
        xls_path = sys.argv[1]
    except ValueError:
        raise ValueError("No file provided")

    if not xls_path.endswith(".xlsx") or xls_path.endswith(".xls"):
        raise ValueError("Unexpected file type. Provided file must be an excel file.")
    
    try:
        output_dir = sys.argv[2]
    except:
        raise ValueError("Output directory not specified.")
    
    table_queue : queue.Queue[Table] = queue.Queue()
    generator = AsyncTableGenerator(table_queue)
    generator.generate_tables([xls_path])

    while generator.is_running():
        poll_table_queue(output_dir)
        sleep(0.1)