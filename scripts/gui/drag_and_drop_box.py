import customtkinter as ctk
from customtkinter import TOP, BOTTOM
from PIL import Image
from tkinterdnd2 import TkinterDnD, DND_ALL

class Tk(ctk.CTk, TkinterDnD.DnDWrapper):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.TkdndVersion = TkinterDnD._require(self)

class DnDBox:
    dnd_icon = ctk.CTkImage(dark_image=Image.open("resources/dnd_white.png"), light_image=Image.open("resources/dnd_black.png"), size=(20, 20))

    def __init__(self, master, on_drop, on_select):
        self.master = master
        self.frame = ctk.CTkFrame(
            master,
            border_width=1
            )
        self.frame.drop_target_register(DND_ALL)  # type: ignore
        self.frame.dnd_bind("<<Drop>>", on_drop) # type: ignore

        # Inner container for centering button + label
        self._inner_container = ctk.CTkFrame(self.frame, fg_color="transparent")

        self._inner_container.update_idletasks()

        # Button for browsing files
        self.select_button = ctk.CTkButton(
            self._inner_container, 
            text="Select files",
            command=on_select,
            width=int(self._inner_container.winfo_reqwidth() * 0.6),
            height=int(self._inner_container.winfo_reqheight() * 0.15),
            font=("Segoe UI", 20, "bold")
        )

        # Drag and drop text label
        self.label = ctk.CTkLabel(
            self._inner_container, 
            image=self.dnd_icon,
            compound="left",
            text=" or drag and drop files here"
            )
        
    def pack_inner(self):
        self._inner_container.pack(expand=True)
        self.select_button.pack(side=TOP, anchor="center", padx=5, pady=5)
        self.label.pack(side=BOTTOM, anchor="center", padx=5, pady=5)
