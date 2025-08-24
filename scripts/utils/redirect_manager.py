from contextlib import contextmanager
import sys

class StringRedirector:
    def __init__(self):
        self.text = ""

    def write(self, s):
        self.text += s

    def flush(self):
        # Needed because some code may call sys.stdout.flush()
        pass

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