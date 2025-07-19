from tkinter import TOP, BOTTOM, LEFT, RIGHT
from tkinterdnd2 import TkinterDnD, DND_ALL
import customtkinter as ctk
from gui.file_item import FileItem
from gui.pop_up_window import PopUpWindow
from table_generation.table import Table
from gui.text_box_redirect import TextboxRedirector
from table_generation.async_table_generator import AsyncTableGenerator
from utils.redirect_manager import redirect_stdout_to
from PIL import Image
from typing import List
import os, sys, platform, queue

ASPECT_RATIO = 9 / 16
RES_X = 720
RES_Y = round(RES_X * ASPECT_RATIO)
RESOLUTION = f"{RES_X}x{RES_Y}"

# Save stdout to restore later
original_stdout = sys.stdout

# Store chosen files
selected_file_paths = set()
file_items = []
recieved_tables : List[Table] = []

# Store frames to hide/show one at a time
frames = []
current_frame = 0

# Object for generating tables asynchronously 
table_queue = queue.Queue()
async_table_generator = AsyncTableGenerator(table_queue)

class Tk(ctk.CTk, TkinterDnD.DnDWrapper):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.TkdndVersion = TkinterDnD._require(self)

def disable_button(button): 
    button.configure(state="disabled", fg_color="#5a5a5a", text_color="gray80")

def display_ui_element(elm, **kwargs):
    elm.pack(**kwargs)

def hide_ui_element(elm):
    elm.pack_forget()
    elm.place_forget()
    elm.grid_forget()

def remove_file_item(file_item : FileItem):
    selected_file_paths.remove(file_item.file_path)

def add_files(paths : List[str]):
    wrong_files = []
    for path in paths:
        if path not in selected_file_paths:
            if path.endswith(".xls") or path.endswith(".xlsx"):
                add_file_item(path)
            else:
                wrong_files.append(path)
    
    if wrong_files:
        show_wrong_file_popup(wrong_files)

def add_file_item(path):
    item = FileItem(file_list_scroll_frame, path, remove_file_item)
    file_items.append(item)
    selected_file_paths.add(path)

def drag_and_drop_files(event):
    raw_data = event.data.strip()
    file_paths = raw_data.split("}")  # supports multiple files
    cleaned_paths = [path.strip("{} ") for path in file_paths]
    next_frame(fill="both")

    add_files(cleaned_paths)

def add_more_files():
    file_paths = ctk.filedialog.askopenfilenames()
    if file_paths:
        add_files(file_paths)

def select_files():
    add_more_files()
    next_frame(fill="both")

def show_wrong_file_popup(file_paths : List[str]):
    popup_win = PopUpWindow(root, "Wrong file type", "Provided files must be excel files.")

    def add_anyway():
        for path in file_paths:
            add_file_item(path)
        popup_win.destroy()

    add_button = ctk.CTkButton(popup_win, text="Add anyway", command=add_anyway)
    ok_button = ctk.CTkButton(popup_win, text="Ok", command=popup_win.destroy)
    add_button.pack(side=LEFT, padx=5, pady=5)
    ok_button.pack(side=RIGHT, padx=5, pady=5)

def next_frame(**kwargs):
    global current_frame

    hide_ui_element(frames[current_frame])
    current_frame += 1

    if current_frame >= len(frames):
        return

    display_ui_element(frames[current_frame], **kwargs)
    back_button.place(x=10, y=10)

def prev_frame(**kwargs):
    global current_frame

    hide_ui_element(frames[current_frame])
    current_frame -= 1

    if current_frame < 0:
        return

    display_ui_element(frames[current_frame], **kwargs)

def back():
    print(current_frame)
    match current_frame:
        case 1:
            hide_ui_element(back_button)
            prev_frame(**FRAME_1_KW)
        case 2:
            prev_frame(**FRAME_2_KW)
            disable_button(save_button)
            async_table_generator.stop_event.set()
        case _:
            return # Back button should not do anything on first frame

def gen_tables():
    global recieved_tables
    
    next_frame(fill="both")
    recieved_tables = [] # Clear tables left from previous generate
    async_table_generator.stop_event.clear() # Make sure the stop flag is set to false

    # Generate tables asynchronously
    try:
        async_table_generator.generate_tables(selected_file_paths)
    except Exception as e:
        print(f"❌ Error: {e}")

    poll_table_queue()


def poll_table_queue():
    try:
        table = table_queue.get_nowait()
        recieved_tables.append(table)
    except queue.Empty:
        # Table generation completed
        if async_table_generator.is_done():
            # Allow saving of files
            save_button.configure(state="normal", fg_color=choose_file_button._fg_color, text_color="white")

        pass
    root.after(100, poll_table_queue)  # Poll every 100ms

def save_tables():
    # Ask user to choose a folder
    output_dir = ctk.filedialog.askdirectory(title="Select output directory")
    if not output_dir:
        return  # User cancelled
    
    # Make subfolders for each selected file only if multiple are selected
    make_subfolders = len(selected_file_paths) > 1

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

    open_button = ctk.CTkButton(confirm_win, text="Open Folder", command=open_folder)
    ok_button = ctk.CTkButton(confirm_win, text="Ok", command=confirm_win.destroy)
    open_button.pack(side=LEFT, padx=5, pady=5)
    ok_button.pack(side=RIGHT, padx=5, pady=5)


if __name__ == "__main__":
    ctk.set_appearance_mode("system")

    # root for window
    root = Tk()
    root.geometry(RESOLUTION)
    root.title("Table Generator")

    # Header
    header_label = ctk.CTkLabel(
        root, 
        text="Table Generator",
        font=("Segoe UI", 50, "bold")
    )
    header_label.pack(side=TOP)

    sub_header_label = ctk.CTkLabel(
        root, 
        text="Generate Word tables from excel data"
    )
    sub_header_label.pack(side=TOP, pady=(0, 10))

    #==================================================
    # Containers (frames) for the different pages
    #==================================================

    # Drag-and-drop frame
    dnd_frame = ctk.CTkFrame(
        root, 
        width=round(RES_X * 0.8), 
        height=round(RES_Y * 0.5),
        border_color="#444",
        border_width=1,
        )
    # Prevent the frame from resizing to fit its contents
    dnd_frame.pack_propagate(False)
    dnd_frame.drop_target_register(DND_ALL)  # type: ignore
    dnd_frame.dnd_bind("<<Drop>>", drag_and_drop_files) # type: ignore

    # Container for elements shown once files have been chosen
    files_chosen_frame = ctk.CTkFrame(root, fg_color="transparent")
    # Frame to display while generating tables
    generating_frame = ctk.CTkFrame(root, fg_color="transparent")

    frames = [dnd_frame, files_chosen_frame, generating_frame]

    #==================================================
    # Defining UI elements and inner containers
    #==================================================
    
    # Back button
    back_img = ctk.CTkImage(light_image=Image.open("resources/back_arrow_white.png"), size=(20, 20))
    back_button = ctk.CTkButton(
        root, 
        image=back_img,
        text="",
        command=back,
        width=30,
    )

    # Button for browsing files
    choose_file_button = ctk.CTkButton(
        dnd_frame, 
        text="Choose Excel files",
        command=select_files,
        width=250,
        height=60,
        font=("Segoe UI", 20, "bold")
    )

    # Drag and drop text label
    dnd_img = ctk.CTkImage(dark_image=Image.open("resources/dnd_white.png"), light_image=Image.open("resources/dnd_black.png"), size=(20, 20))
    dnd_file_label = ctk.CTkLabel(
        dnd_frame, 
        image=dnd_img,
        compound="left",
        text=" or drag and drop files here"
        )
    
    # File list container
    file_list_frame = ctk.CTkFrame(files_chosen_frame, width=round(RES_X * 0.8))
    
    # Scrollable frame for file list
    file_list_scroll_frame = ctk.CTkScrollableFrame(file_list_frame, fg_color="transparent", height=130)
    file_list_scroll_frame._scrollbar.configure(height=130) # Internal minimum is 200 by default, change to 130

    # Add files button
    add_files_img = ctk.CTkImage(dark_image=Image.open("resources/add_files_white.png"), light_image=Image.open("resources/add_files_black.png"), size=(20, 20))

    more_files_button = ctk.CTkButton(
        file_list_frame, 
        image=add_files_img,
        compound="left",
        text=" Add more files",
        command=add_more_files
    )

    # Button for generating the tables
    generate_button = ctk.CTkButton(
        files_chosen_frame,
        text="Generate tables",
        width=250,
        height=60,
        font=("Segoe UI", 20, "bold"),
        command=gen_tables
    )

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
        state="disabled",
        fg_color="#5a5a5a",        # Gray background
        text_color="gray80", 
        command=save_tables
    ) 

    #==================================================
    # Placing UI elements and inner containers
    #==================================================

    # Frame 1
    FRAME_1_KW = {"side":TOP, "padx":10, "pady":10, "expand":True, "fill":"both"}
    dnd_frame.pack(**FRAME_1_KW)
    choose_file_button.place(relx=0.5, rely=0.3, anchor="center")
    dnd_file_label.place(relx=0.5, rely=0.5, anchor="center")
    
    # Frame 2
    FRAME_2_KW = {"fill":"both", "expand":True, "padx":10, "pady":10}
    file_list_frame.pack(**FRAME_2_KW)
    file_list_scroll_frame.pack(fill="both", expand=True, pady=5)
    more_files_button.pack(side=LEFT, padx=5, pady=5)
    generate_button.pack(side=TOP, anchor="e", padx=5, pady=10)

    # Frame 3
    output_textbox.pack(fill="both", padx=10, pady=10)
    save_button.pack(side=TOP, anchor="e", padx=10, pady=(5,10))

    root.mainloop()
