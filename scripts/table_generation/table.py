from docx.document import Document
from dataclasses import dataclass
import os

@dataclass
class Table:
    doc : Document
    source_path : str
    component_id : str

    def save(self, output_dir, make_subfolder=True):
        if make_subfolder:
            subfolder_name = os.path.basename(self.source_path).replace(" ", "_")
            full_subfolder_path = os.path.join(output_dir, subfolder_name)
        else:
            full_subfolder_path = output_dir

        if not os.path.isdir(full_subfolder_path):
            os.makedirs(full_subfolder_path)

        save_path = os.path.join(full_subfolder_path, f"table_{self.component_id}.docx")
        save_path = os.path.normpath(save_path)
        print(f"Saved table in {save_path}")
        self.doc.save(save_path)