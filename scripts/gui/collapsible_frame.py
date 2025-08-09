import customtkinter as ctk

class CollapsibleFrame(ctk.CTkFrame):
    def __init__(self, master, title="Section", expanded=False, *args, **kwargs):
        super().__init__(master, *args, **kwargs)

        self.expanded = expanded

        # Toggle button
        self.toggle_button = ctk.CTkButton(self, text=f"{'▼' if self.expanded else '►'} {title}",
                                           command=self.toggle, anchor="w")
        self.toggle_button.pack(fill="x", padx=5, pady=5)

        # Frame that holds the collapsible content
        self.content = ctk.CTkFrame(self)
        self.content.pack(fill="x", expand=False, padx=5, pady=(0, 5))

        # Start collapsed or expanded
        if not self.expanded:
            self.content.pack_forget()

    def toggle(self):
        self.expanded = not self.expanded
        if self.expanded:
            self.content.pack(fill="x", expand=False, padx=5, pady=(0, 5))
            self.toggle_button.configure(text=f"▼ {self.toggle_button.cget('text')[2:]}")
        else:
            self.content.pack_forget()
            self.toggle_button.configure(text=f"► {self.toggle_button.cget('text')[2:]}")

    def add_widget(self, widget):
        widget.pack(in_=self.content, fill="x", pady=2)

