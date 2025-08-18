import customtkinter as ctk
from customtkinter import ThemeManager
from word_sync.sync_files import Mismatch
from utils.gui_utils import blend_colors, get_color
from typing import Callable
from PIL import Image

def _get_similarity_color(similarity: float) -> str:
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

def _diff_words(a: str, b: str):
    words_a = a.split()
    words_b = b.split()
    
    result = []
    match_run = []
    mismatch_run_a = []
    mismatch_run_b = []

    # Iterate over both lists of words
    for wa, wb in zip(words_a, words_b):
        if wa == wb:
            # Flush any accumulated mismatches
            if mismatch_run_a:
                result.append((" ".join(mismatch_run_a), " ".join(mismatch_run_b)))
                mismatch_run_a = []
                mismatch_run_b = []
            match_run.append(wa)
        else:
            # Flush any accumulated matches
            if match_run:
                result.append(" ".join(match_run))
                match_run = []
            mismatch_run_a.append(wa)
            mismatch_run_b.append(wb)

    # Flush any remaining runs
    if match_run:
        result.append(" ".join(match_run))
    if mismatch_run_a:
        result.append((" ".join(mismatch_run_a), " ".join(mismatch_run_b)))

    # Handle leftover words
    if len(words_a) > len(words_b):
        result.append((" ".join(words_a[len(words_b):]), None))
    elif len(words_b) > len(words_a):
        result.append((None, " ".join(words_b[len(words_a):])))

    return result

class DifferenceFrame(ctk.CTkFrame):
    def __init__(self, master, text: str, button_text, on_press, icon_path=None, icon_text="", **kwargs):
        super().__init__(master, **kwargs)

        if icon_path:
            img = ctk.CTkImage(Image.open(icon_path))
            self._label = ctk.CTkLabel(self, text=icon_text, image=img, compound="left", text_color="gray")
            self._label.pack(anchor="w", padx=5, pady=(5, 0))

        font = ctk.CTkFont(family="Arial", size=16)
        self.label = ctk.CTkLabel(self, text=text, font=font, anchor="w")

        self.label.pack(expand=True, padx=5, pady=5)
        ctk.CTkButton(self, text=button_text, command=on_press).pack(fill="x", expand=True, padx=5, pady=5)

        self.bind("<Configure>", self._update_wraplength)

    def _update_wraplength(self, event):
        # Make wraplength match current label width
        self.label.configure(wraplength=self.winfo_width()/ctk.ScalingTracker.get_window_scaling(self))

class MismatchItem(ctk.CTkFrame):
    def __init__(
            self, 
            master, 
            mismatch : Mismatch, 
            on_resolve : Callable[['MismatchItem'], None] | None = None,
            **kwargs
            ):
        super().__init__(master, **kwargs)
        self.mismatch = mismatch
        self.on_resolve = on_resolve

        default_fg_color = get_color(self, "CTkFrame", "fg_color")
        default_border_color = get_color(self, "CTkFrame", "border_color")
        similarity_color = _get_similarity_color(mismatch.similarity / 100)
        border_color = blend_colors(similarity_color, default_border_color, 0.4)
        fg_color = blend_colors(border_color, default_fg_color, 0.2)

        self.configure(border_color=border_color, fg_color=fg_color, border_width=2)

        # Variable to store the result
        self.result_var = ctk.StringVar(value="")

        self.grid_columnconfigure((0, 1), weight=1, uniform="row2")
        self.grid_rowconfigure(0, weight=0)
        self.grid_rowconfigure(1, weight=0)
        self.grid_rowconfigure(2, weight=1)
        self.grid_rowconfigure(3, weight=0)

        from gui import MultiPartLabel

        bold_font = ctk.CTkFont(family="Seoge UI", size=18, weight="bold")
        font = ctk.CTkFont(family="Seoge UI", size=18)
        match_text = "mismatch ❌" if int(mismatch.similarity) != 100 else "match ✅"
        MultiPartLabel(
            self, 
            parts=[
                {
                    "text":f"{mismatch.mismatch_type.upper()} ",
                    "font":bold_font,
                    "text_color":border_color
                },
                {
                    "text":match_text,
                    "font":font,
                    "text_color":border_color
                },
            ]
            ).grid(row=0, column=0, padx=3, pady=5, sticky="nw")

        ctk.CTkLabel(
            self, 
            text=f"{int(mismatch.similarity)}% match",
            font=font,
            text_color=border_color
            ).grid(row=0, column=1, padx=5, pady=5, sticky="ne")
        
        header_text = mismatch.header if mismatch.mismatch_type == "description" else ""
        ctk.CTkLabel(
            self, 
            text=f"in {header_text}",
            font=font,
            text_color=border_color
            ).grid(row=1, column=0, padx=5, pady=5, sticky="nw")

        if int(mismatch.similarity) != 100:
            print(_diff_words(mismatch.in_word, mismatch.in_excel))

            w_frame = DifferenceFrame(
                self, 
                text=f"{mismatch.in_word}", 
                button_text="Use Word", 
                on_press=self._button_cmd("w"),
                icon_path="resources/word_icon.png",
                icon_text=" In Word",
                corner_radius=2
                )
            e_frame = DifferenceFrame(
                self, 
                text=f"{mismatch.in_excel}", 
                button_text="Use Excel",
                on_press=self._button_cmd("e"),
                icon_path="resources/excel_icon.png",
                icon_text=" In Excel",
                corner_radius=2
                )
            skip_button = ctk.CTkButton(self, text="Skip", command=self._button_cmd("s"))

            w_frame.grid(row=2, column=0, padx=5, pady=5, sticky="nsew")
            e_frame.grid(row=2, column=1, padx=5, pady=5, sticky="nsew")
            skip_button.grid(row=3, column=1, padx=5, pady=5, sticky="se")

    def get_choice(self):
        """Block until a choice is made and return it."""
        try:
            self.wait_variable(self.result_var)  # freezes until set
            return self.result_var.get()
        except Exception as e:
            print(e)

    def resolve(self):
        if self.on_resolve:
            self.on_resolve(self)
    
    def _button_cmd(self, res_str : str):
        def f():
            self.result_var.set(res_str)
            self.resolve()
        return f
    
class MismatchContainer(ctk.CTkScrollableFrame):
    def __init__(self, master, **kwargs):
        super().__init__(master, **kwargs)
        self._last = None

    def add_mismatch(self, mismatch : Mismatch, **kwargs):

        def on_resolve(mismatch_item : MismatchItem):
            pass
            
        mismatch_item = MismatchItem(
            self, 
            mismatch,
            on_resolve=on_resolve,
            )

        if self._last:
            self._last.resolve()
        self._last = mismatch_item
        mismatch_item.pack(**kwargs)

        # Wait until Tkinter finishes rendering
        self.after_idle(self._scroll_to_bottom)

    def _scroll_to_bottom(self):
        self._parent_canvas.yview_moveto(1)

    def get_choice(self):
        if self._last:
            return self._last.get_choice()
        return "s" # skip command