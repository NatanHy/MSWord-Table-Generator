import customtkinter as ctk

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

    def set_part_text(self, index, text):
        """Change text of a specific part"""
        if 0 <= index < len(self.labels):
            self.labels[index].configure(text=text)

    def set_part_color(self, index, color):
        """Change text color of a specific part"""
        if 0 <= index < len(self.labels):
            self.labels[index].configure(text_color=color)

    def set_part_font(self, index, font):
        """Change font of a specific part"""
        if 0 <= index < len(self.labels):
            self.labels[index].configure(font=font)