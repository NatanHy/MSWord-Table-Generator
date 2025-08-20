import customtkinter as ctk
from customtkinter import ThemeManager

class ProgressBar(ctk.CTkFrame):
    def __init__(self, parent, progress_var: ctk.DoubleVar, width=400, height=15, **kwargs):
        super().__init__(parent, width=width, height=height + 20, **kwargs)  # extra height for labels
        self.progress_var = progress_var

        # Theme colors
        self.bg_color = ThemeManager.theme["CTkFrame"]["fg_color"]
        self.fg_color = ThemeManager.theme["CTkButton"]["fg_color"]

        # Top-right percentage label
        self.perc_label = ctk.CTkLabel(self, text="0%")
        self.perc_label.pack(side="right", padx=5)

        # Progress bar background
        self.bar_frame = ctk.CTkFrame(self, width=width, height=height, fg_color=self.bg_color, corner_radius=height//3)
        self.bar_frame.pack(fill="x", expand=True)

        # Foreground progress bar
        self.fg_bar = ctk.CTkFrame(self.bar_frame, fg_color=self.fg_color, corner_radius=height//3)
        self.fg_bar.place(relx=0, rely=0, relwidth=0, relheight=1)

        # Trace variables
        self.progress_var.trace_add("write", self._update_progress)
        self._update_progress()  # initialize


    def _update_progress(self, *args):
        value = self.progress_var.get()
        value = max(0.0, min(1.0, value))
        self.fg_bar.place_configure(relwidth=value)
        self.perc_label.configure(text=f"{int(value * 100)}%")