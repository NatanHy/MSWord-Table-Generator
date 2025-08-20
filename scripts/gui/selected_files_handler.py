from typing import Iterable

from gui.file_item import FileItem
import customtkinter as ctk
from tkinterdnd2 import DND_ALL

class _UI:
    def __init__(self, master, **kwargs):
        self.initialized = False
        if master is not None:
            # Force height to 0 to allow the frame to be as small as needed, by default the minimum height is 200
            self.file_list_scroll_frame = ctk.CTkScrollableFrame(master, **kwargs)
            self.file_list_scroll_frame._scrollbar.configure(height=0) 
            self.initialized = True

    def pack(self, **kwargs):
        if not self.initialized:
            raise RuntimeError("UI not initialized.")
        if self is not None:
            self.file_list_scroll_frame.pack(**kwargs)

    def grid(self, **kwargs):
        if not self.initialized:
            raise RuntimeError("UI not initialized.")
        if self is not None:
            self.file_list_scroll_frame.grid(**kwargs)

    def configure(self, **kwargs):
        if not self.initialized:
            raise RuntimeError("UI not initialized.")
        self.file_list_scroll_frame.configure(**kwargs)

class SelectedFilesHandler:
    def __init__(self, master=None, filter=lambda s: True, on_wrong=None, after_add=None):
        self.filter = filter
        self.selected_file_paths : set[str] = set()
        self.file_items = []
        self.on_wrong = on_wrong
        self.after_add = after_add

        self.add_ui(master)

    @property
    def has_files(self):
        return len(self.selected_file_paths) > 0

    @property
    def first_path(self):
        return next(iter(self.selected_file_paths))

    def add_ui(self, master, **kwargs):
        self.ui = _UI(master, **kwargs)
        if master is not None:
            self.ui.file_list_scroll_frame._parent_frame.drop_target_register(DND_ALL) # type: ignore
            self.ui.file_list_scroll_frame._parent_frame.dnd_bind("<<Drop>>", self.drag_and_drop_files) # type: ignore

    def add_files(self, file_paths : Iterable[str]):
        for f in file_paths:
            self._add_file(f)
        if self.after_add is not None:
            self.after_add()

    def drag_and_drop_files(self, event):
        raw_data = event.data.strip()
        file_paths = raw_data.split("}")  # supports multiple files
        cleaned_paths = [path.strip("{} ") for path in file_paths]

        self.add_files(cleaned_paths)

    def select_files(self):
        file_paths = ctk.filedialog.askopenfilenames()
        if file_paths:
            self.add_files(file_paths) 

    def _add_file(self, file_path : str):
        if file_path in self.selected_file_paths:
            return
        
        wrong_files = []

        if self.filter(file_path):
            self.selected_file_paths.add(file_path)
            self._add_file_item(file_path)
        else:
            wrong_files.append(file_path)
        
        if self.on_wrong is not None and wrong_files:
            self.on_wrong(self, wrong_files)

    def _remove_file_item(self, file_item : FileItem):
        self.selected_file_paths.remove(file_item.file_path)
        
    def _add_file_item(self, path):
        if self.ui.initialized:
            item = FileItem(self.ui.file_list_scroll_frame, path, self._remove_file_item)
            item.pack(fill="x", padx=0, pady=1)
            self.file_items.append(item)
