from tkinter import TOP, LEFT, RIGHT
import customtkinter as ctk
from gui.file_item import FileItem
from gui.pop_up_window import PopUpWindow
from table_generation.table import TableCollection
from gui.text_box_redirect import TextboxRedirector
from table_generation.async_table_generator import AsyncTableGenerator
from utils.redirect_manager import redirect_stdout_to
from PIL import Image
from typing import List, Iterable
import os, sys, platform, queue
from docx import Document
import docx.document
from tkinterdnd2 import TkinterDnD, DND_ALL
from gui.drag_and_drop_box import DnDBox
from customtkinter import ThemeManager
from gui.collapsible_frame import CollapsibleFrame
from gui.selected_files_handler import SelectedFilesHandler

ASPECT_RATIO = 9 / 16
RES_X = 720
RES_Y = round(RES_X * ASPECT_RATIO)
RESOLUTION = f"{RES_X}x{RES_Y}"

# Save stdout to restore later
original_stdout = sys.stdout

# Store tables/documents
recieved_tables : List[TableCollection] = []
doc_path_for_insertion : str = ""
doc_for_insertion : docx.document.Document | None = None

# Store frames to hide/show one at a time
frames = []
current_frame = 0

# Object for generating tables asynchronously 
table_queue = queue.Queue()
async_table_generator = AsyncTableGenerator(table_queue)

def disable_button(button): 
    if button._state != "disabled":
        button.configure(state="disabled", fg_color="#5a5a5a", text_color="gray80")

def enable_button(button):
    if button._state != "normal":
        clr = ThemeManager.theme["CTkButton"]["fg_color"]
        button.configure(state="normal", fg_color=clr, text_color="white")

def display_ui_element(elm, **kwargs):
    elm.pack(**kwargs)

def hide_ui_element(elm):
    elm.pack_forget()
    elm.place_forget()
    elm.grid_forget()

def show_wrong_file_popup(file_handler : SelectedFilesHandler, file_paths : List[str]):
    popup_win = PopUpWindow(root, "Wrong file type", "Wrong file type.")

    def add_anyway():
        prev_filter = file_handler.filter
        file_handler.filter = lambda s: True # Disable filter
        file_handler.add_files(file_paths)
        file_handler.filter = prev_filter # Restore filter
        popup_win.destroy()

    add_button = ctk.CTkButton(popup_win, text="Add anyway", command=add_anyway)
    ok_button = ctk.CTkButton(popup_win, text="Ok", command=popup_win.destroy)
    add_button.pack(side=LEFT, padx=5, pady=5)
    ok_button.pack(side=RIGHT, padx=5, pady=5)

def go_to_frame(frame, **kwargs):
    global current_frame
    hide_ui_element(frames[current_frame])

    current_frame = frame

    display_ui_element(frames[current_frame], **kwargs)

    if current_frame != 0:
        back_button.place(x=10, y=10)

def prev_frame(**kwargs):
    global current_frame

    hide_ui_element(frames[current_frame])
    current_frame -= 1

    if current_frame < 0:
        return

    display_ui_element(frames[current_frame], **kwargs)

def back():
    match current_frame:
        case 1:
            hide_ui_element(back_button)
            prev_frame(**FRAME_0_KW)
        case 2:
            prev_frame(**FRAME_1_KW)
        case 3:
            prev_frame(**FRAME_2_KW)
            disable_button(save_button)
            async_table_generator.stop_event.set()
        case _:
            return # Back button should not do anything on first frame

def gen_tables(insert=False):
    global recieved_tables, doc_for_insertion, doc_path_for_insertion
    
    go_to_frame(3, fill="both")
    recieved_tables = [] # Clear tables left from previous generate
    async_table_generator.stop_event.clear() # Make sure the stop flag is set to false

    # Generate tables asynchronously
    try:
        if insert:
            doc_path_for_insertion = insert_doc_file_handler.first_path
            doc_for_insertion = Document(doc_path_for_insertion)
            async_table_generator.generate_and_insert_tables(excel_file_handler.selected_file_paths, doc_for_insertion)
        else:
            async_table_generator.template_file_path = empty_doc_file_handler.first_path
            async_table_generator.generate_tables(excel_file_handler.selected_file_paths)
    except:
        pass

    # When inserting into a document the queue is not used, but this function also
    # enables the save button once the generation is done
    poll_table_queue() 

def poll_table_queue():
    try:
        table = table_queue.get_nowait()
        recieved_tables.append(table)
    except queue.Empty:
        # Table generation completed
        if async_table_generator.is_done():
            # Allow saving of files
            enable_button(save_button)

        pass
    root.after(100, poll_table_queue)  # Poll every 100ms

def save_tables():
    # If we are inserting into a word file, just save the file and return
    if doc_for_insertion is not None:
        doc_for_insertion.save(doc_path_for_insertion)
        show_save_confirmation(os.path.dirname(doc_path_for_insertion))
        return

    # Otherwise prompt user for a folder and save there
    output_dir = ctk.filedialog.askdirectory(title="Select output directory")
    if not output_dir:
        return  # User cancelled
    
    # Make subfolders for each selected file only if multiple are selected
    make_subfolders = len(excel_file_handler.selected_file_paths) > 1

    with redirect_stdout_to(output_redirector):
        for table in recieved_tables:
            table.save(output_dir, make_subfolder=make_subfolders)

    # Show confirmation popup with "Open folder" option
    show_save_confirmation(output_dir)

def show_save_confirmation(folder_path):
    confirm_win = PopUpWindow(root, "Save Complete", "✅ Tables saved successfully!")

    def open_folder():
        if platform.system() == "Windows":
            os.startfile(folder_path)
        elif platform.system() == "Darwin":  # macOS
            os.system(f"open '{folder_path}'")
        else:  # Linux
            os.system(f"xdg-open '{folder_path}'")

        confirm_win.destroy()

    confirm_win.set_left("Open Folder", open_folder)
    confirm_win.set_right("Ok", confirm_win.destroy)

class Tk(ctk.CTk, TkinterDnD.DnDWrapper):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.TkdndVersion = TkinterDnD._require(self)

if __name__ == "__main__":
    # File handlers
    excel_file_handler = SelectedFilesHandler(
        filter=lambda s: s.endswith(".xlsx"), 
        on_wrong=show_wrong_file_popup,
        after_add=lambda: go_to_frame(1, **FRAME_1_KW)
        )
    empty_doc_file_handler = SelectedFilesHandler(
        filter=lambda s: s.endswith(".docx"), 
        on_wrong=show_wrong_file_popup
        )
    insert_doc_file_handler = SelectedFilesHandler(
        filter=lambda s: s.endswith(".docx"), 
        on_wrong=show_wrong_file_popup,
        )

    ctk.set_appearance_mode("system")

    # root for window
    root = Tk()
    root.geometry(RESOLUTION)
    root.title("Table Generator")

    #==================================================
    # Containers (frames) for the different pages
    #==================================================

    # Drag-and-drop box
    dnd_box = DnDBox(root, on_drop=excel_file_handler.drag_and_drop_files, on_select=excel_file_handler.select_files)
    dnd_box.frame.configure(
        width=round(RES_X * 0.8), 
        height=round(RES_Y * 0.5)
    )
    dnd_box.select_button.configure(
        text="Select Excel Files"
    )
    # Prevent the frame from resizing to fit its contents
    dnd_box.frame.pack_propagate(False)

    # Container for elements shown once files have been chosen
    files_chosen_frame = ctk.CTkFrame(root, fg_color="transparent")
    files_chosen_frame.drop_target_register(DND_ALL) # type: ignore
    files_chosen_frame.dnd_bind("<<Drop>>", excel_file_handler.drag_and_drop_files) # type: ignore

    # Container for choosing generation type/settings
    generate_settings_frame = ctk.CTkFrame(root, fg_color="transparent")

    # Frame to display while generating tables
    generating_frame = ctk.CTkFrame(root, fg_color="transparent")

    frames = [dnd_box.frame, files_chosen_frame, generate_settings_frame, generating_frame]

    #==================================================
    # Defining UI elements and inner containers
    #==================================================
    
    # Header
    header_label = ctk.CTkLabel(
        root, 
        text="Table Generator",
        font=("Segoe UI", 50, "bold")
    )

    sub_header_label = ctk.CTkLabel(
        root, 
        text="Generate Word tables from excel data"
    )

    # Back button
    back_img = ctk.CTkImage(light_image=Image.open("resources/back_arrow_white.png"), size=(20, 20))
    back_button = ctk.CTkButton(
        root, 
        image=back_img,
        text="",
        command=back,
        width=30,
    )
    
    # File list container
    file_list_frame = ctk.CTkFrame(files_chosen_frame, width=round(RES_X * 0.8))
    excel_file_handler.add_ui(file_list_frame, height=0)
    excel_file_handler.ui.configure(fg_color="transparent")
    
    # Add files button
    add_files_img = ctk.CTkImage(dark_image=Image.open("resources/add_files_white.png"), light_image=Image.open("resources/add_files_black.png"), size=(20, 20))

    more_files_button = ctk.CTkButton(
        file_list_frame, 
        image=add_files_img,
        compound="left",
        text=" Add more files",
        command=excel_file_handler.select_files
    )

    # Button for proceeding with table generation
    continue_button = ctk.CTkButton(
        files_chosen_frame,
        text="Continue",
        width=250,
        height=60,
        font=("Segoe UI", 20, "bold"),
        command=lambda: go_to_frame(2, **FRAME_2_KW)
    )

    # Containers for the two generation types
    empty_doc_frame = CollapsibleFrame(generate_settings_frame, title="Generate in empty document", fg_color="transparent")
    insert_doc_frame = CollapsibleFrame(generate_settings_frame, title="Insert into existing document", fg_color="transparent")

    for container in [empty_doc_frame.content, insert_doc_frame.content]:
        # (1, 1) has weight 0 since the generate buttons need to always be visible
        container.grid_columnconfigure(0, weight=1) 
        container.grid_columnconfigure(1, weight=0)
        container.grid_rowconfigure(0, weight=1)
        container.grid_rowconfigure(1, weight=0)

    empty_dnd_box = DnDBox(
        empty_doc_frame.content, 
        on_drop=empty_doc_file_handler.drag_and_drop_files, 
        on_select=empty_doc_file_handler.select_files
        )
    empty_dnd_box.select_button.configure(
        text="Select File (optional)"
    )
    insert_dnd_box = DnDBox(
        insert_doc_frame.content, 
        on_drop=insert_doc_file_handler.drag_and_drop_files, 
        on_select=insert_doc_file_handler.select_files
        )
    insert_dnd_box.select_button.configure(
        text="Select File"
    )

    empty_doc_file_handler.add_ui(empty_doc_frame.content, height=110)
    insert_doc_file_handler.add_ui(insert_doc_frame.content, height=110)

    gen_empty_button = ctk.CTkButton(
        empty_doc_frame.content,
        text="Generate",
        width=250,
        height=60,
        font=("Segoe UI", 20, "bold"),
        command=lambda: gen_tables(insert=False)
    )
    gen_insert_button = ctk.CTkButton(
        insert_doc_frame.content,
        text="Generate",
        width=250,
        height=60,
        font=("Segoe UI", 20, "bold"),
        command=lambda: gen_tables(insert=True)
    )
    disable_button(gen_insert_button)
    insert_doc_file_handler.after_add = lambda: enable_button(gen_insert_button)

    # Command-line style textbox for table generation output
    output_textbox = ctk.CTkTextbox(
        generating_frame, 
        height=200,
        font=("Seoge UI Mono", 12)
        )
    
    output_redirector = TextboxRedirector(output_textbox)
    async_table_generator.stdout_redirect = output_redirector
    
    # Button for saving tables
    save_button = ctk.CTkButton(
        generating_frame,
        text="Save",
        width=250,
        height=60,
        font=("Segoe UI", 20, "bold"),
        command=save_tables
    ) 
    disable_button(save_button)

    #==================================================
    # Placing UI elements and inner containers
    #==================================================

    header_label.pack(side=TOP)
    sub_header_label.pack(side=TOP, pady=(0, 10))

    # Frame 0
    FRAME_0_KW = {"side":TOP, "padx":10, "pady":10, "expand":True, "fill":"both"}
    dnd_box.frame.pack(**FRAME_0_KW)
    dnd_box.pack_inner()
    
    # Frame 1
    FRAME_1_KW = {"fill":"both", "expand":True, "padx":10, "pady":10}
    file_list_frame.pack(**FRAME_1_KW)
    excel_file_handler.ui.pack(fill="both", expand=True, pady=5) #type: ignore
    more_files_button.pack(side=LEFT, padx=5, pady=5)
    continue_button.pack(side=TOP, anchor="e", padx=10, pady=(5, 10))

    # Frame 2
    FRAME_2_KW = {"fill":"both", "expand":True, "padx":10, "pady":(10,5)}
    empty_doc_frame.pack(anchor="n", fill="x", padx=5)
    insert_doc_frame.pack(anchor="n", fill="x", padx=5)
    empty_dnd_box.pack_inner()
    insert_dnd_box.pack_inner()

    empty_dnd_box.frame.grid(row=0, column=0, rowspan=2, sticky="nsew", padx=10, pady=10)
    insert_dnd_box.frame.grid(row=0, column=0, rowspan=2, sticky="nsew", padx=10, pady=10)

    empty_doc_file_handler.ui.grid(row=0, column=1, sticky="nsew", padx=10, pady=(10, 5)) #type: ignore
    insert_doc_file_handler.ui.grid(row=0, column=1, sticky="nsew", padx=10, pady=(10, 5)) #type: ignore

    gen_empty_button.grid(row=1, column=1, sticky="nsew", padx=10, pady=(5, 10))
    gen_insert_button.grid(row=1, column=1, sticky="nsew", padx=10, pady=(5, 10))

    # Frame 3
    output_textbox.pack(fill="both", padx=10, pady=10)
    save_button.pack(side=TOP, anchor="e", padx=10, pady=(5,10))

    root.mainloop()
