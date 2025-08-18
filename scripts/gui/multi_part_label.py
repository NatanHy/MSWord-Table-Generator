import customtkinter as ctk

class MultiPartTextBox(ctk.CTkTextbox):
    def __init__(self, master, parts, **kwargs):
        super().__init__(master, wrap="word", **kwargs)

        # Make it behave like a label
        self.configure(state="normal", fg_color="transparent", border_width=0)

        # Insert styled parts
        for i, part in enumerate(parts):
            tag = f"part{i}"
            self.insert("end", part["text"], tag)
            self.tag_config(tag, **{k:v for k, v in part.items() if k != "text"})

        # Disable editing
        self.configure(state="disabled")

class MultiPartLabel(ctk.CTkFrame):
    def __init__(self, master, parts, **kwargs):
        """
        parts: list of dicts, each dict describing a text segment:
            {
                "text": str,
                "font": ctk.CTkFont or None,
                "text_color": str or None
            }
        Example:
            [
                {"text": "Word: ", "font": bold_font, "text_color": "black"},
                {"text": "John Smith", "text_color": "red"}
            ]
        """
        super().__init__(master, **kwargs)

        self.configure(fg_color="transparent")

        self.labels = []
        for i, part in enumerate(parts):
            lbl = ctk.CTkLabel(
                self,
                **part
            )
            lbl.pack(side="left", padx=0, pady=0)
            self.labels.append(lbl)

