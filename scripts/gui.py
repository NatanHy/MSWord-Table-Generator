from tkinter import TOP, BOTTOM, LEFT, RIGHT
from tkinterdnd2 import TkinterDnD, DND_ALL
import customtkinter as ctk
from file_item import FileItem
from text_box_redirect import TextboxRedirector
from PIL import Image
from main import generate_tables
from pathlib import Path
import os, sys, threading, platform

ASPECT_RATIO = 9 / 16
RES_X = 720
RES_Y = round(RES_X * ASPECT_RATIO)
RESOLUTION = f"{RES_X}x{RES_Y}"

# Save stdout to restore later
original_stdout = sys.stdout

selected_file_paths = set()
file_items = []
per_file_tables = {}

class Tk(ctk.CTk, TkinterDnD.DnDWrapper):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.TkdndVersion = TkinterDnD._require(self)

def display_ui_element(elm, **kwargs):
    elm.pack(**kwargs)

def hide_ui_element(elm):
    elm.pack_forget()

def remove_file_item(file_item : FileItem):
    selected_file_paths.remove(file_item.file_path)

def add_files(paths):
    for path in paths:
        if path not in selected_file_paths:
            item = FileItem(file_list_scroll_frame, path, remove_file_item)
            file_items.append(item)
            selected_file_paths.add(path)

def drag_and_drop_files(event):
    raw_data = event.data.strip()
    file_paths = raw_data.split("}")  # supports multiple files
    cleaned_paths = [path.strip("{} ") for path in file_paths]
    add_files(cleaned_paths)

    show_file_display()

def select_files():
    file_paths = ctk.filedialog.askopenfilenames()
    add_files(file_paths)

    show_file_display()

# Called to display slected files once user has chosen one or more files
def show_file_display():
    hide_ui_element(dnd_frame)
    display_ui_element(files_chosen_frame, fill="both")

def show_generate_display():
    hide_ui_element(files_chosen_frame)
    display_ui_element(generating_frame, fill="both")

def run_generate_tables_async():
    threading.Thread(target=_gen_tables, daemon=True).start()

def _gen_tables():
    global original_stdout
    sys.stdout = TextboxRedirector(output_textbox)  # Redirect print to GUI

    try:
        for path in selected_file_paths:
            tables = generate_tables(path)
            path_name = Path(path).stem
            per_file_tables[path_name] = tables
    except Exception as e:
        print(f"❌ Error: {e}")
    finally:
        sys.stdout = original_stdout  # Restore normal stdout
        save_button.configure(state="normal", fg_color=choose_file_button._fg_color, text_color="white") # Allow saving of files

def gen_tables():
    show_generate_display()
    run_generate_tables_async()

def save_tables():
    # Ask user to choose a folder
    output_dir = ctk.filedialog.askdirectory(title="Select output directory")
    if not output_dir:
        return  # User cancelled

    items = per_file_tables.items()
    for path, tables in items:
        # Sanitize subdirectory name

        if len(items) > 1:
            subfolder_name = os.path.basename(path).replace(" ", "_")
            full_subfolder_path = os.path.join(output_dir, subfolder_name)
        else:
            full_subfolder_path = output_dir

        if not os.path.isdir(full_subfolder_path):
            os.makedirs(full_subfolder_path)

        for table_name, document in tables.items():
            save_path = os.path.join(full_subfolder_path, table_name)
            document.save(save_path)

    # Show confirmation popup with "Open folder" option
    show_save_confirmation(output_dir)

def show_save_confirmation(folder_path):
    confirm_win = ctk.CTkToplevel()
    confirm_win.title("Save Complete")

    # --- Center the popup over the main window ---
    confirm_win.update_idletasks()
    main_x = root.winfo_rootx()
    main_y = root.winfo_rooty()
    main_width = root.winfo_width()
    main_height = root.winfo_height()

    popup_width = 300
    popup_height = 120

    pos_x = main_x + (main_width // 2) - (popup_width // 2)
    pos_y = main_y + (main_height // 2) - (popup_height // 2)

    confirm_win.geometry(f"{popup_width}x{popup_height}+{pos_x}+{pos_y}")
    confirm_win.resizable(False, False)

    confirm_win.lift()           # Bring to front
    confirm_win.focus_force()    # Grab focus
    confirm_win.attributes("-topmost", True)  # Force it above all windows

    label = ctk.CTkLabel(confirm_win, text="✅ Tables saved successfully!")
    label.pack(pady=(20, 5))

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

    # Drag-and-drop frame
    dnd_frame = ctk.CTkFrame(
        root, 
        width=round(RES_X * 0.8), 
        height=round(RES_Y * 0.5),
        border_color="#444",
        border_width=1,
        )

    dnd_frame.pack(side=TOP, padx=10, pady=10, expand=True, fill="both")
    dnd_frame.drop_target_register(DND_ALL)  # type: ignore
    dnd_frame.dnd_bind("<<Drop>>", drag_and_drop_files) # type: ignore

    # Prevent the frame from resizing to fit its contents
    dnd_frame.pack_propagate(False)

    # Button for browsing files
    choose_file_button = ctk.CTkButton(
        dnd_frame, 
        text="Choose files",
        command=select_files,
        width=250,
        height=60,
        font=("Segoe UI", 20, "bold")
    )
    choose_file_button.place(relx=0.5, rely=0.3, anchor="center")

    # Drag and drop text label
    dnd_img = ctk.CTkImage(dark_image=Image.open("resources/dnd_white.png"), light_image=Image.open("resources/dnd_black.png"), size=(20, 20))

    dnd_file_label = ctk.CTkLabel(
        dnd_frame, 
        image=dnd_img,
        compound="left",
        text=" or drag and drop files here"
        )
    dnd_file_label.place(relx=0.5, rely=0.5, anchor="center")

    # Container for elements shown once files have been chosen
    files_chosen_frame = ctk.CTkFrame(root, fg_color="transparent")

    # File list container
    file_list_frame = ctk.CTkFrame(files_chosen_frame, width=round(RES_X * 0.8))
    file_list_frame.pack(fill="both", expand=True, padx=5, pady=(10, 0))

    file_list_scroll_frame = ctk.CTkScrollableFrame(file_list_frame, fg_color="transparent", height=130)
    file_list_scroll_frame._scrollbar.configure(height=130) # Internal minimum is 200 by default, change to 130
    file_list_scroll_frame.pack(fill="both", expand=True, pady=5)

    # Add files button
    add_files_img = ctk.CTkImage(dark_image=Image.open("resources/add_files_white.png"), light_image=Image.open("resources/add_files_black.png"), size=(20, 20))

    more_files_button = ctk.CTkButton(
        file_list_frame, 
        image=add_files_img,
        compound="left",
        text=" Add more files",
        command=select_files
    )
    more_files_button.pack(side=LEFT, padx=5, pady=5)

    # Generate button below, aligned right
    generate_button = ctk.CTkButton(
        files_chosen_frame,
        text="Generate tables",
        width=250,
        height=60,
        font=("Segoe UI", 20, "bold"),
        command=gen_tables
    )
    generate_button.pack(side=TOP, anchor="e", padx=5, pady=10)

    # Frame to display while generating tables
    generating_frame = ctk.CTkFrame(root, fg_color="transparent")

    # Command-line style textbox for table generation output
    output_textbox = ctk.CTkTextbox(
        generating_frame, 
        height=200,
        font=("Seoge UI Mono", 12)
        )
    output_textbox.pack(fill="both", padx=10, pady=10)

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
    save_button.pack(side=TOP, anchor="e", padx=10, pady=(5,10))

    root.mainloop()
