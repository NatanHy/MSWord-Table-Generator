import os
import queue
import sys
from typing import List
import traceback

import customtkinter as ctk
from customtkinter import LEFT, TOP, ThemeManager
from docx import Document
import docx.document
from PIL import Image

from gui import (
    CollapsibleFrame, 
    DnDBox, 
    PopUpWindow, 
    SelectedFilesHandler, 
    TextboxRedirector,  
    OnHover, 
    FrameManager
    )
from table_generation.table import TableCollection
from table_generation.async_table_generator import AsyncTableGenerator
from utils.gui_utils import (
    disable_button, 
    open_folder, 
    wrong_files_popup, 
    color_filter, 
    disable_button_while,
    switch_theme
    )
from utils.redirect_manager import redirect_stdout_to
from utils.files import create_backup, resource_path

class App(ctk.CTkFrame):
    ASPECT_RATIO = 9 / 16
    RES_X = 720
    RES_Y = round(RES_X * ASPECT_RATIO)
    RESOLUTION = f"{RES_X}x{RES_Y}"

    FRAME_0_KW = {"side":TOP, "padx":10, "pady":10, "expand":True, "fill":"both"}
    FRAME_1_KW = {"fill":"both", "expand":True, "padx":5, "pady":10}
    FRAME_2_KW = {"fill":"both", "expand":True, "padx":10, "pady":(10,5)}
    FRAME_KWARGS = {
        0: {"side":TOP, "padx":10, "pady":10, "expand":True, "fill":"both"},
        1: {"fill":"both", "expand":True, "padx":5, "pady":10},
        2: {"fill":"both", "expand":True, "padx":10, "pady":(10,5)},
        3: {"fill":"both"}
    }

    def __init__(self, master):
        super().__init__(master)
        self._fg_color = "transparent"
        
        # Save stdout to restore later
        self.original_stdout = sys.stdout

        # Store tables/documents
        self.recieved_tables : List[TableCollection] = []
        self.doc_path_for_insertion : str = ""
        self.doc_for_insertion : docx.document.Document | None = None

        # Object for generating tables asynchronously 
        self.table_queue = queue.Queue()
        self.async_table_generator = AsyncTableGenerator(self.table_queue)

    # Callback when using back button on frame 3 specifically
    def _back_3(self):
        disable_button(self.save_button)
        self.async_table_generator.stop_event.set()

    def _gen_tables(self, insert=False):        
        self.frame_manager.go_to_frame(3)
        self.recieved_tables = [] # Clear tables left from previous generate
        self.async_table_generator.stop_event.clear() # Make sure the stop flag is set to false
        self.async_table_generator.on_fail = self._show_gen_fail

        # Generate tables asynchronously
        try:
            if insert:
                assert self.insert_doc_file_handler.has_files, "No files for insertion found"
                self.doc_path_for_insertion = self.insert_doc_file_handler.first_path() #type: ignore
                self.doc_for_insertion = Document(self.doc_path_for_insertion)
                self.async_table_generator.generate_and_insert_tables(self.excel_file_handler.selected_file_paths, self.doc_for_insertion)
            else:
                self.async_table_generator.template_file_path = self.empty_doc_file_handler.first_path()
                self.async_table_generator.generate_tables(self.excel_file_handler.selected_file_paths)
        except:
            pass

        # When inserting into a document the queue is not used, but this function also
        # enables the save button once the generation is done
        self._poll_table_queue() 

    def _poll_table_queue(self):
        try:
            table = self.table_queue.get_nowait()
            self.recieved_tables.append(table)
        except queue.Empty:
            pass
        self.after(100, self._poll_table_queue)  # Poll every 100ms

    def _save_tables(self):
        try:
            # If we are inserting into a word file, just save the file and return
            if self.doc_for_insertion is not None:
                create_backup(self.doc_path_for_insertion)
                self.doc_for_insertion.save(self.doc_path_for_insertion)
                self._show_save_confirmation(os.path.dirname(self.doc_path_for_insertion))
                return

            # Otherwise prompt user for a folder and save there
            output_dir = ctk.filedialog.askdirectory(title="Select output directory")
            if not output_dir:
                return  # User cancelled
            
            # Make subfolders for each selected file only if multiple are selected
            make_subfolders = len(self.excel_file_handler.selected_file_paths) > 1

            with redirect_stdout_to(self.output_redirector):
                for table in self.recieved_tables:
                    table.save(output_dir, make_subfolder=make_subfolders)

            # Show confirmation popup with "Open folder" option
            self._show_save_confirmation(output_dir)
        except Exception as e:
            print(traceback.format_exc())
            self._show_save_fail(e)

    def _show_save_confirmation(self, folder_path):
        confirm_win = PopUpWindow(self, "Save Complete", "✅ Tables saved successfully!")

        def open_and_destroy():
            open_folder(folder_path)
            confirm_win.destroy()

        confirm_win.set_left("Open Folder", open_and_destroy)
        confirm_win.set_right("Ok", confirm_win.destroy)

    def _show_save_fail(self, err):
        confirm_win = PopUpWindow(self, "Save Failed", f"Could not save files:\n{err}")
        confirm_win.set_left("Cancel", confirm_win.destroy)
        confirm_win.set_right("Ok", confirm_win.destroy)

    def _show_gen_fail(self, err : Exception):
        confirm_win = PopUpWindow(self, "Generation Failed", f"Could not generate tables:\n{err}")

        # Function to copy traceback
        def copy_traceback():
            self.clipboard_clear()
            tb_str = traceback.format_exception(type(err), err, err.__traceback__)
            self.clipboard_append(tb_str) #type: ignore
            self.update()  # Needed to update clipboard
            confirm_win.label.configure(text="Traceback copied to clipboard")

        confirm_win.set_left("Copy traceback", copy_traceback)
        confirm_win.set_right("Ok", confirm_win.destroy)

    def stop(self):
        self.destroy()

    def run(self):
        # File handlers
        self.excel_file_handler = SelectedFilesHandler(
            filter=lambda s: s.endswith(".xlsx"), 
            on_wrong=wrong_files_popup(self, "Wrong file type, file must be Excel file (.xlsx)")
            )
        self.empty_doc_file_handler = SelectedFilesHandler(
            filter=lambda s: s.endswith(".docx"), 
            on_wrong=wrong_files_popup(self,  "Wrong file type, file must be Word file (.docx)")
            )
        self.insert_doc_file_handler = SelectedFilesHandler(
            filter=lambda s: s.endswith(".docx"), 
            on_wrong=wrong_files_popup(self, "Wrong file type, file must be Word file (.docx)")
            )

        if not os.path.exists("backups"):
            os.makedirs("backups", exist_ok=True)

        #==================================================
        # Containers (frames) for the different pages
        #==================================================

        header_frame = ctk.CTkFrame(self, fg_color="transparent")

        # Drag-and-drop box
        dnd_box = DnDBox(
            self, 
            on_drop=self.excel_file_handler.drag_and_drop_files, 
            on_select=self.excel_file_handler.select_files,
            )
        dnd_box.frame.configure(
            width=round(self.RES_X * 0.8), 
            height=round(self.RES_Y * 0.5)
        )
        dnd_box.select_button.configure(
            text="Select Excel Files"
        )
        # Prevent the frame from resizing to fit its contents
        dnd_box.frame.pack_propagate(False)

        # Container for elements shown once files have been chosen
        files_chosen_frame = ctk.CTkFrame(self, fg_color="transparent")

        # Container for choosing generation type/settings
        generate_settings_frame = ctk.CTkFrame(self, fg_color="transparent")

        # Frame to display while generating tables
        generating_frame = ctk.CTkFrame(self, fg_color="transparent")

        # Frame manager for chaning between frames
        self.frame_manager = FrameManager(
            self, 
            frames=[dnd_box.frame, files_chosen_frame, generate_settings_frame, generating_frame],
            frame_kwargs=self.FRAME_KWARGS,
            on_back_callbacks={3: self._back_3}
            )

        self.excel_file_handler.after_add=lambda: self.frame_manager.go_to_frame(1)

        #==================================================
        # Defining UI elements and inner containers
        #==================================================
        
        self.header_container = ctk.CTkFrame(header_frame, fg_color="transparent")

        # Header
        self.header_label = ctk.CTkLabel(
            self.header_container, 
            text="Table Generator",
            font=("Segoe UI", 50, "bold")
        )

        self.sub_header_label = ctk.CTkLabel(
            self.header_container, 
            text="Generate Word tables from excel data"
        )
        
        # Backup button
        white_image = Image.open(resource_path("resources/back_up_white.png"))
        colored_image = color_filter(white_image, ThemeManager.theme["CTkButton"]["fg_color"])
        backup_img = ctk.CTkImage(light_image=colored_image, size=(20, 20))
        
        self.backup_button = ctk.CTkButton(
            header_frame, 
            image=backup_img,
            text="",
            fg_color="transparent",
            border_color=ThemeManager.theme["CTkButton"]["fg_color"],
            border_width=1,
            command=lambda: open_folder("backups"),
            width=30,
        )
        # save object instance to stop python's garbage collector form deleting it
        _hover1 = OnHover(self.backup_button, "Open backups folder")

        # Button for changing Light/Dark theme
        sun = Image.open(resource_path("resources/sun.png"))
        colored_sun_image = color_filter(sun, ThemeManager.theme["CTkButton"]["fg_color"])
        moon = Image.open(resource_path("resources/moon.png"))
        colored_moon_image = color_filter(moon, ThemeManager.theme["CTkButton"]["fg_color"])
        theme_change_img = ctk.CTkImage(light_image=colored_moon_image, dark_image=colored_sun_image, size=(20, 20))
        
        self.theme_change_button = ctk.CTkButton(
            header_frame, 
            image=theme_change_img,
            text="",
            fg_color="transparent",
            border_color=ThemeManager.theme["CTkButton"]["fg_color"],
            border_width=1,
            command=switch_theme,
            width=30,
        )    
        _hover2 = OnHover(self.theme_change_button, "Change theme")


        # File list container
        file_list_frame = ctk.CTkFrame(files_chosen_frame, width=round(self.RES_X * 0.8))
        self.excel_file_handler.add_ui(file_list_frame, height=0)
        self.excel_file_handler.ui.configure(fg_color="transparent")
        
        # Add files button
        add_files_img = ctk.CTkImage(light_image=Image.open(resource_path("resources/add_files_white.png")), size=(20, 20))

        self.more_files_button = ctk.CTkButton(
            file_list_frame, 
            image=add_files_img,
            compound="left",
            text=" Add more files",
            command=self.excel_file_handler.select_files
        )

        # Button for proceeding with table generation
        self.continue_button = ctk.CTkButton(
            files_chosen_frame,
            text="Continue",
            width=250,
            height=60,
            font=("Segoe UI", 20, "bold"),
            command=lambda: self.frame_manager.go_to_frame(2)
        )

        # Containers for the two generation types
        self.empty_doc_frame = CollapsibleFrame(generate_settings_frame, title="Generate in empty document", fg_color="transparent")
        self.insert_doc_frame = CollapsibleFrame(generate_settings_frame, title="Insert into existing document", fg_color="transparent")

        self.empty_dnd_box = DnDBox(
            self.empty_doc_frame.content, 
            on_drop=self.empty_doc_file_handler.drag_and_drop_files, 
            on_select=self.empty_doc_file_handler.select_files
            )
        self.empty_dnd_box.select_button.configure(
            text="Select Template (optional)"
        )
        _hover3 = OnHover(self.empty_dnd_box.select_button, "Select optional template document for styling. The document will not be modified")

        self.insert_dnd_box = DnDBox(
            self.insert_doc_frame.content, 
            on_drop=self.insert_doc_file_handler.drag_and_drop_files, 
            on_select=self.insert_doc_file_handler.select_files
            )
        self.insert_dnd_box.select_button.configure(
            text="Select Document"
        )

        self.empty_doc_file_handler.add_ui(self.empty_doc_frame.content, height=110)
        self.insert_doc_file_handler.add_ui(self.insert_doc_frame.content, height=110)

        self.gen_empty_button = ctk.CTkButton(
            self.empty_doc_frame.content,
            text="Generate",
            width=250,
            height=60,
            font=("Segoe UI", 20, "bold"),
            command=lambda: self._gen_tables(insert=False)
        )

        def _disable_gen_while() -> bool:
            return not self.insert_doc_file_handler.has_files
        self.gen_insert_button = ctk.CTkButton(
            self.insert_doc_frame.content,
            text="Generate",
            width=250,
            height=60,
            font=("Segoe UI", 20, "bold"),
            command=lambda: self._gen_tables(insert=True)
        )
        disable_button_while(self.gen_insert_button, _disable_gen_while)

        # Command-line style textbox for table generation output
        output_textbox = ctk.CTkTextbox(
            generating_frame, 
            height=200,
            font=("Seoge UI Mono", 12)
            )
        
        self.output_redirector = TextboxRedirector(output_textbox)
        self.async_table_generator.stdout_redirect = self.output_redirector
        
        def _disable_save_while() -> bool:
            return not self.async_table_generator.is_done()

        # Button for saving tables
        self.save_button = ctk.CTkButton(
            generating_frame,
            text="Save",
            width=250,
            height=60,
            font=("Segoe UI", 20, "bold"),
            command=self._save_tables
        ) 
        disable_button_while(self.save_button, _disable_save_while)

        #==================================================
        # Placing UI elements and inner containers
        #==================================================

        header_frame.pack(fill="x")
        self.header_container.pack(fill="both")

        self.header_label.pack()
        self.sub_header_label.pack()
        self.backup_button.place(relx=1.0, anchor="ne", x=-10, y=10)
        self.theme_change_button.place(relx=1.0, anchor="ne", x=-50, y=10)

        # Frame 0
        dnd_box.frame.pack(**self.FRAME_0_KW)
        dnd_box.pack_inner()
        
        # Frame 1
        file_list_frame.pack(**self.FRAME_1_KW)
        self.excel_file_handler.ui.pack(fill="both", expand=True, pady=5) #type: ignore
        self.more_files_button.pack(side=LEFT, padx=5, pady=5)
        self.continue_button.pack(side=TOP, anchor="e", padx=5, pady=5)

        # Frame 2
        self.empty_doc_frame.pack(anchor="n", fill="x", padx=5)
        self.insert_doc_frame.pack(anchor="n", fill="x", padx=5)
        self.empty_dnd_box.pack_inner()
        self.insert_dnd_box.pack_inner()

        for container in [self.empty_doc_frame.content, self.insert_doc_frame.content]:
            # (1, 1) has weight 0 since the generate buttons need to always be visible
            container.grid_columnconfigure(0, weight=1) 
            container.grid_columnconfigure(1, weight=0)
            container.grid_rowconfigure(0, weight=1)
            container.grid_rowconfigure(1, weight=0)

        self.empty_dnd_box.frame.grid(row=0, column=0, rowspan=2, sticky="nsew", padx=10, pady=10)
        self.insert_dnd_box.frame.grid(row=0, column=0, rowspan=2, sticky="nsew", padx=10, pady=10)

        self.empty_doc_file_handler.ui.grid(row=0, column=1, sticky="nsew", padx=10, pady=(10, 5)) #type: ignore
        self.insert_doc_file_handler.ui.grid(row=0, column=1, sticky="nsew", padx=10, pady=(10, 5)) #type: ignore

        self.gen_empty_button.grid(row=1, column=1, sticky="nsew", padx=10, pady=(5, 10))
        self.gen_insert_button.grid(row=1, column=1, sticky="nsew", padx=10, pady=(5, 10))

        # Frame 3
        output_textbox.pack(fill="both", padx=10, pady=10)
        self.save_button.pack(side=TOP, anchor="e", padx=10, pady=(5,10))