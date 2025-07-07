class TextboxRedirector:
    def __init__(self, textbox):
        self.textbox = textbox

    def write(self, message):
        # Schedule update on main thread
        self.textbox.after(0, self._write, message)

    def _write(self, message):
        self.textbox.insert("end", message)
        self.textbox.see("end")  # Auto-scroll

    def flush(self):
        pass  # Needed for compatibility