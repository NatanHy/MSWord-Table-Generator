import customtkinter as ctk
from customtkinter import LEFT, RIGHT

class PopUpWindow(ctk.CTkToplevel):
    def __init__(self, master, title, text, width=300, height=120):
        super().__init__(master)
        self.label = ctk.CTkLabel(self, text=text, wraplength=width)

        self.resizable(True, False)
        self.pack_propagate(True)
        self.title(title)

        self.label.pack(pady=(20, 5))
        # Delay geometry call to allow rendering
        self.after(10, lambda: self._center(master, width, height))

    def set_left(self, text, cmd):
        button = ctk.CTkButton(self, text=text, command=cmd)
        button.pack(side=LEFT, padx=5, pady=10)

        self.update_idletasks()

    def set_right(self, text, cmd):
        button = ctk.CTkButton(self, text=text, command=cmd)
        button.pack(side=RIGHT, padx=5, pady=10)

        self.update_idletasks()

    def _center(self, master, width, height):
        master.update_idletasks()

        # Account for scaling in width/height since geoemtry() takes in absolute pixel coordinates
        scaling = self.tk.call('tk', 'scaling')
        main_x = master.winfo_rootx()
        main_y = master.winfo_rooty()
        main_width = master.winfo_width() / scaling
        main_height = master.winfo_height() / scaling

        pos_x = int(main_x + (main_width / 2) - (width / 2))
        pos_y = int(main_y + (main_height / 2) - (height / 2))

        self.geometry(f"{pos_x}+{pos_y}")

        # Bring to front
        self.lift()
        self.focus_force()
        self.attributes("-topmost", True)