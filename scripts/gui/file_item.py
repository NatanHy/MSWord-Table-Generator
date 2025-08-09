import customtkinter as ctk
import os

class FileItem(ctk.CTkFrame):
    def __init__(self, master, file_path, on_remove):
        super().__init__(
            master,            
            fg_color="transparent", 
            border_color="#444",    
            border_width=1,
            corner_radius=6
            )
        self.file_path = file_path
        self.remove_callback = on_remove

        self.grid_columnconfigure(0, weight=1)
        self.grid_columnconfigure(1, weight=0) 

        # Use grid to make sure remove button stays on the right side for long 
        self.label = ctk.CTkLabel(self, text=os.path.basename(file_path), anchor="w")
        self.label.grid(row=0, column=0, sticky="nsew", padx=(10, 5), pady=5)

        self.remove_button = ctk.CTkButton(self, text="X", width=30, command=self.remove, fg_color="#ff0000")
        self.remove_button.grid(row=0, column=1, sticky="e", padx=(5, 10), pady=5)
        self.pack(fill="x")

    def remove(self):
        self.destroy()
        self.remove_callback(self)