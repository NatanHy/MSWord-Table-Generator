import customtkinter as ctk
from word_sync.sync_files import Mismatch

class MismatchItem(ctk.CTkFrame):
    def __init__(self, master, mismatch, **kwargs):
        super().__init__(master, **kwargs)
        self.mismatch = mismatch

        border_color = self._get_similarity_color(mismatch.similarity)
        self.configure(border_color=border_color, border_width=1)

        # Variable to store the result
        self.result_var = ctk.StringVar(value="")

        ctk.CTkLabel(self, text=f"{mismatch.mismatch_type} mismatch found in {mismatch.header}").pack()
        ctk.CTkLabel(self, text=f"Word: {mismatch.in_word}").pack()
        ctk.CTkLabel(self, text=f"Excel: {mismatch.in_excel}").pack()

        ctk.CTkButton(self, text="Keep Word", command=lambda: self.result_var.set("w")).pack(side="left", expand=True)
        ctk.CTkButton(self, text="Keep Excel", command=lambda: self.result_var.set("e")).pack(side="left", expand=True)
        ctk.CTkButton(self, text="Skip", command=lambda: self.result_var.set("s")).pack(side="left", expand=True)

    def get_choice(self):
        """Block until a choice is made and return it."""
        self.wait_variable(self.result_var)  # freezes until set
        return self.result_var.get()

    def _get_similarity_color(self, similarity: float) -> str:
        """
        Map similarity (0-1) to a color from dark red to green.
        Returns a hex string.
        """
        # Clamp similarity
        s = max(0.0, min(1.0, similarity))
        # Linear interpolation: red -> green
        r = int(255 * (1 - s))
        g = int(255 * s)
        b = 100  # constant to keep it darker
        return f"#{r:02x}{g:02x}{b:02x}"


    # --- Handlers that call the provided callbacks ---
    def _handle_keep_word(self):
        if self.on_keep_word:
            self.on_keep_word()

    def _handle_keep_excel(self):
        if self.on_keep_excel:
            self.on_keep_excel()

    def _handle_skip(self):
        if self.on_skip:
            self.on_skip()