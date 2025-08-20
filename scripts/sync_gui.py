import os

import customtkinter as ctk
from customtkinter import ThemeManager, BOTTOM
from PIL import Image

from gui import Tk, OnHover, SelectedFilesHandler, FrameManager, MismatchContainer, PopUpWindow
from word_sync import WordExcelSyncer
from utils.gui_utils import open_folder, wrong_files_popup, color_filter, disable_button_while

ASPECT_RATIO = 9 / 16
RES_X = 720
RES_Y = round(RES_X * ASPECT_RATIO)
RESOLUTION = f"{RES_X}x{RES_Y}"

FRAME_0_KW = FRAME_1_KW = {"fill":"both", "expand":True}
FRAME_KWARGS = {0:FRAME_0_KW, 1:FRAME_1_KW}

sync_done = False

def save_files():
    try:
        file_syncer.save_files()
        folder_path = os.path.dirname(word_file_handler.first_path)
        show_save_confirmation(folder_path)
    except Exception as e:
        show_save_fail(e)

def show_save_fail(err):
    confirm_win = PopUpWindow(root, "Save Failed", f"Could not save synced files:\n{err}")
    confirm_win.set_left("Cancel", confirm_win.destroy)
    confirm_win.set_right("Ok", confirm_win.destroy)

def show_save_confirmation(folder_path):
    confirm_win = PopUpWindow(root, "Save Complete", "✅ Synced files saved successfully!")

    def open_and_destroy():
        open_folder(folder_path)
        confirm_win.destroy()

    confirm_win.set_left("Open Folder", open_and_destroy)
    confirm_win.set_right("Ok", confirm_win.destroy)

def set_sync_done_false():
    global sync_done
    sync_done = False

def sync():
    global sync_done
    original_protocol = root.protocol("WM_DELETE_WINDOW")

    try:
        doc_path = word_file_handler.first_path
        xls_paths = list(excel_file_handler.selected_file_paths)
        frame_manager.go_to_frame(1)
        frame_manager.frames[frame_manager.current_frame].update_idletasks()
        gen = file_syncer.sync_files(doc_path, xls_paths)

        # Set WM_DELETE_WINDOW protocol to break the while loop
        def stop_loop():
            mismatch_container._last.result_var.set("EXIT") #type: ignore
        root.protocol("WM_DELETE_WINDOW", stop_loop)

        while True:
            mismatch = next(gen)  # Start the generator

            mismatch_container.update_idletasks()
            mismatch_container.add_mismatch(mismatch, fill="both", expand=True, padx=5, pady=5)

            if int(mismatch.similarity) != 100:            
                decision = mismatch_container.get_choice()
            else:
                decision = "s"

            if decision == "EXIT":
                break
            mismatch = gen.send(decision)
        root.destroy() # Only reached when WM_DELETE_WINODW is called
    except StopIteration:
        # restore WM_DELETE_WINDOW protocol
        root.protocol("WM_DELETE_WINDOW", lambda: root.tk.call(original_protocol))
        sync_done = True


if __name__ == "__main__":
    ctk.set_appearance_mode("system")

    root = Tk()
    root.geometry(RESOLUTION)
    root.title("File syncing")
    
    word_file_handler = SelectedFilesHandler(
        filter=lambda s: s.endswith(".docx"),
        on_wrong=wrong_files_popup(root, "Wrong file type, file must be Word file (.docx)")
        )
    excel_file_handler = SelectedFilesHandler(
        filter=lambda s: s.endswith(".xlsx"),
        on_wrong=wrong_files_popup(root, "Wrong file type, file must be Excel file (.xlsx)")
        )
    file_syncer = WordExcelSyncer()

    #==================================================
    # Defining UI elements and inner containers
    #==================================================

    selection_frame = ctk.CTkFrame(root, fg_color="transparent")
    syncing_frame = ctk.CTkFrame(root, fg_color="transparent")

    frame_manager = FrameManager(
        root,
        frames=[selection_frame, syncing_frame],
        frame_kwargs=FRAME_KWARGS,
        on_back_callbacks={1:set_sync_done_false},
        back_button_pos=(5, 5)
    )
    
    # Backup button
    white_image = Image.open("resources/back_up_white.png")
    colored_image = color_filter(white_image, ThemeManager.theme["CTkButton"]["fg_color"])
    backup_img = ctk.CTkImage(light_image=colored_image, size=(20, 20))
    
    backup_button = ctk.CTkButton(
        root, 
        image=backup_img,
        text="",
        fg_color="transparent",
        border_color=ThemeManager.theme["CTkButton"]["fg_color"],
        border_width=1,
        command=lambda: open_folder("backups"),
        width=30,
    )
    # save object instance to stop python's garbage collector form deleting it
    _hover = OnHover(backup_button, "Open backups folder")

    select_word_button = ctk.CTkButton(
        selection_frame, 
        text="Select Word document", 
        command=word_file_handler.select_files,
        height=50,
        font=("Segoe UI", 16, "bold")
        )
    select_excel_button = ctk.CTkButton(
        selection_frame, 
        text="Select Excel file", 
        command=excel_file_handler.select_files,
        height=50,
        font=("Segoe UI", 16, "bold")
        )

    def _disable_sync_while() -> bool:
        # Keep sync button disabled while missing word or excel files
        has_files = word_file_handler.has_files and excel_file_handler.has_files
        return not has_files
    
    sync_img = ctk.CTkImage(light_image=Image.open("resources/sync.png"), size=(40, 40)) 
    sync_button = ctk.CTkButton(
        selection_frame, 
        text="Sync", 
        image=sync_img,
        command=sync,
        height=50,
        width=200,
        font=("Segoe UI", 20, "bold")
        )
    disable_button_while(sync_button, _disable_sync_while)

    word_file_handler.add_ui(selection_frame)
    excel_file_handler.add_ui(selection_frame)

    mismatch_container = MismatchContainer(syncing_frame)

    # Button for saving synced files
    save_button = ctk.CTkButton(
        syncing_frame,
        text="Save",
        width=250,
        height=60,
        font=("Segoe UI", 20, "bold"),
        command=save_files
    ) 
    disable_button_while(save_button, lambda: not sync_done)

    #==================================================
    # Placing UI elements and inner containers
    #==================================================

    backup_button.pack(anchor="e", padx=5, pady=(5, 0))

    selection_frame.pack(**FRAME_0_KW)
    selection_frame.grid_columnconfigure((0, 1), weight=1)
    selection_frame.grid_rowconfigure((0, 1, 3), weight=0)
    selection_frame.grid_rowconfigure(2, weight=1) # Only let the file display expand


    select_word_button.grid(row=1, column=0, sticky="ew", padx=5, pady=5)
    select_excel_button.grid(row=1, column=1, sticky="ew", padx=5, pady=5)
    word_file_handler.ui.grid(row=2, column=0, sticky="nsew", padx=5, pady=5)
    excel_file_handler.ui.grid(row=2, column=1, sticky="nsew", padx=5, pady=5)
    sync_button.grid(row=3, column=1, sticky="nse", padx=5, pady=(5, 10))

    mismatch_container.pack(fill="both", expand=True, padx=5, pady=5)
    save_button.pack(side=BOTTOM, anchor="e", padx=5, pady=5)

    root.mainloop()

