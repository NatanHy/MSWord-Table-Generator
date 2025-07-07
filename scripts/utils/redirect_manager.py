from contextlib import contextmanager
import sys

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