import customtkinter as ctk
from utils.gui_utils import *
from gui import Tk, OnHover, SelectedFilesHandler, FrameManager
from gui.mismatch_item import MismatchItem
from PIL import Image
from word_sync import WordExcelSyncer

ASPECT_RATIO = 9 / 16
RES_X = 720
RES_Y = round(RES_X * ASPECT_RATIO)
RESOLUTION = f"{RES_X}x{RES_Y}"

FRAME_0_KW = {"fill":"both", "expand":True}
FRAME_KWARGS = {0:FRAME_0_KW}

def _similarity_to_rgb(score):
    t = score / 100

    r = int((1 - t) * 255)
    g = int(t * t * 255)
    b = 0

    return r, g, b

def _colored_similarity_text(score):
    r, g, b = _similarity_to_rgb(score)
    return f"\033[38;2;{r};{g};{b}m({score:.2f}% similar)\033[0m"

def sync():
    try:
        doc_path = word_file_handler.first_path
        xls_paths = list(excel_file_handler.selected_file_paths)
        frame_manager.go_to_frame(1)

        gen = file_syncer.sync_files(doc_path, xls_paths)
        mismatch = next(gen)  # Start the generator
        while True:
            # sim_text = _colored_similarity_text(mismatch.similarity)
            # print(f"\033[31mFound {mismatch.mismatch_type} mismatch \033[0m{sim_text}:")
            # print(f"\tIn '{mismatch.header}'")
            # print(f"\tWord:  '{mismatch.in_word}'")
            # print(f"\tExcel: '{mismatch.in_excel}'")

            mismatch_item = MismatchItem(
                syncing_frame, 
                mismatch
                )
            
            mismatch_item.pack()
            decision = mismatch_item.get_choice()
            mismatch = gen.send(decision)
    except StopIteration:
        pass


if __name__ == "__main__":
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
        frame_kwargs=FRAME_KWARGS
    )
    
    # Backup button
    white_image = Image.open("resources/back_up_white.png")
    colored_image = color_filter(white_image, ThemeManager.theme["CTkButton"]["fg_color"])
    backup_img = ctk.CTkImage(light_image=colored_image, size=(20, 20))
    
    backup_button = ctk.CTkButton(
        selection_frame, 
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


    #==================================================
    # Placing UI elements and inner containers
    #==================================================

    selection_frame.pack(**FRAME_0_KW)
    selection_frame.grid_columnconfigure((0, 1), weight=1)
    selection_frame.grid_rowconfigure((0, 1, 3), weight=0)
    selection_frame.grid_rowconfigure(2, weight=1) # Only let the file display expand

    backup_button.grid(row=0, column=1, sticky="ne", padx=5, pady=(5, 0))

    select_word_button.grid(row=1, column=0, sticky="ew", padx=5, pady=5)
    select_excel_button.grid(row=1, column=1, sticky="ew", padx=5, pady=5)
    word_file_handler.ui.grid(row=2, column=0, sticky="nsew", padx=5, pady=5)
    excel_file_handler.ui.grid(row=2, column=1, sticky="nsew", padx=5, pady=5)
    sync_button.grid(row=3, column=1, sticky="nse", padx=5, pady=(5, 10))

    root.mainloop()

