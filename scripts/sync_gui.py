import os

import customtkinter as ctk
from customtkinter import ThemeManager, BOTTOM, RIGHT
from PIL import Image

from gui import (
    Tk, 
    OnHover, 
    SelectedFilesHandler, 
    FrameManager, 
    MismatchContainer, 
    PopUpWindow,
    ProgressBar
    )
from word_sync import WordExcelSyncer
from utils.gui_utils import (
    open_folder, 
    wrong_files_popup, 
    color_filter, 
    disable_button_while, 
    switch_theme
    )
from utils.files import resource_path
from utils.redirect_manager import redirect_stdout_to, StringRedirector

class App(ctk.CTkFrame):
    ASPECT_RATIO = 9 / 16
    RES_X = 720
    RES_Y = round(RES_X * ASPECT_RATIO)
    RESOLUTION = f"{RES_X}x{RES_Y}"

    FRAME_0_KW = FRAME_1_KW = {"fill":"both", "expand":True, "pady":0}
    FRAME_KWARGS = {0:FRAME_0_KW, 1:FRAME_1_KW}

    def __init__(self, master : Tk):
        super().__init__(master)
        self.master = master
        self._fg_color = "transparent"

        self.sync_done = False

    def save_files(self):
        try:
            self.file_syncer.save_files()
            folder_path = os.path.dirname(self.word_file_handler.first_path()) #type: ignore
            self.show_save_confirmation(folder_path)
        except Exception as e:
            self.show_save_fail(e)

    def show_save_fail(self, err):
        confirm_win = PopUpWindow(self, "Save Failed", f"Could not save synced files:\n{err}")
        confirm_win.set_left("Cancel", confirm_win.destroy)
        confirm_win.set_right("Ok", confirm_win.destroy)

    def show_save_confirmation(self, folder_path):
        confirm_win = PopUpWindow(self, "Save Complete", "✅ Synced files saved successfully!")

        def open_and_destroy():
            open_folder(folder_path)
            confirm_win.destroy()

        confirm_win.set_left("Open Folder", open_and_destroy)
        confirm_win.set_right("Ok", confirm_win.destroy)

    def show_err_log(self):
        confirm_win = PopUpWindow(self, "Error log", text=self.error_log.text, width=600, text_box=True)

    def set_sync_done_false(self):
        self.sync_done = False

    def sync(self):
        assert isinstance(self.master, Tk)
        original_protocol = self.master.protocol("WM_DELETE_WINDOW")

        with redirect_stdout_to(self.error_log):
            try:
                doc_path = self.word_file_handler.first_path()
                xls_paths = list(self.excel_file_handler.selected_file_paths)
                self.frame_manager.go_to_frame(1)
                self.frame_manager.frames[self.frame_manager.current_frame].update_idletasks()
                gen = self.file_syncer.sync_files(doc_path, xls_paths, progress_var=self.progress_var) #type: ignore

                # Set WM_DELETE_WINDOW protocol to break the while loop
                def stop_loop():
                    self.mismatch_container._last.result_var.set("EXIT") #type: ignore
                self.master.protocol("WM_DELETE_WINDOW", stop_loop)

                mismatch = next(gen)  # Start the generator
                while True:
                    self.mismatch_container.update_idletasks()
                    self.mismatch_container.add_mismatch(mismatch, fill="both", expand=True, padx=5, pady=5)

                    if int(mismatch.similarity) != 100:            
                        decision =self. mismatch_container.get_choice()
                    else:
                        decision = "s"

                    if decision == "EXIT":
                        break
                    mismatch = gen.send(decision)

                # Only reached when WM_DELETE_WINODW is called
                self.master.protocol("WM_DELETE_WINDOW", lambda: self.master.tk.call(original_protocol))
                self.master.tk.call(original_protocol)
            except StopIteration:
                # restore WM_DELETE_WINDOW protocol
                self.master.protocol("WM_DELETE_WINDOW", lambda: self.master.tk.call(original_protocol))
                self.sync_done = True

    def _disable_sync_while(self) -> bool:
        # Keep sync button disabled while missing word or excel files
        has_files = self.word_file_handler.has_files and self.excel_file_handler.has_files
        return not has_files

    def run(self):
        self.word_file_handler = SelectedFilesHandler(
            filter=lambda s: s.endswith(".docx"),
            on_wrong=wrong_files_popup(self, "Wrong file type, file must be Word file (.docx)")
            )
        self.excel_file_handler = SelectedFilesHandler(
            filter=lambda s: s.endswith(".xlsx"),
            on_wrong=wrong_files_popup(self, "Wrong file type, file must be Excel file (.xlsx)")
            )
        self.file_syncer = WordExcelSyncer()

        if not os.path.exists("backups"):
            os.makedirs("backups", exist_ok=True)

        #==================================================
        # Defining UI elements and inner containers
        #==================================================

        header_frame = ctk.CTkFrame(self, fg_color="transparent")
        selection_frame = ctk.CTkFrame(self, fg_color="transparent")
        syncing_frame = ctk.CTkFrame(self, fg_color="transparent")

        self.frame_manager = FrameManager(
            self,
            frames=[selection_frame, syncing_frame],
            frame_kwargs=self.FRAME_KWARGS,
            on_back_callbacks={1:self.set_sync_done_false},
            back_button_pos=(5, 5)
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
        _hover = OnHover(self.backup_button, "Open backups folder")

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

        self.select_word_button = ctk.CTkButton(
            selection_frame, 
            text="Select Word document", 
            command=self.word_file_handler.select_files,
            height=50,
            font=("Segoe UI", 16, "bold")
            )
        self.select_excel_button = ctk.CTkButton(
            selection_frame, 
            text="Select Excel files", 
            command=self.excel_file_handler.select_files,
            height=50,
            font=("Segoe UI", 16, "bold")
            )
        
        sync_img = ctk.CTkImage(light_image=Image.open(resource_path("resources/sync.png")), size=(40, 40)) 
        self.sync_button = ctk.CTkButton(
            selection_frame, 
            text="Sync", 
            image=sync_img,
            command=self.sync,
            height=50,
            width=200,
            font=("Segoe UI", 20, "bold")
            )
        disable_button_while(self.sync_button, self._disable_sync_while)

        self.word_file_handler.add_ui(selection_frame)
        self.excel_file_handler.add_ui(selection_frame)

        self.mismatch_container = MismatchContainer(syncing_frame)

        self.progress_var = ctk.DoubleVar(value=0.0)
        self.progress_bar = ProgressBar(syncing_frame, self.progress_var, fg_color="transparent")

        # Button for saving synced files
        self.save_button = ctk.CTkButton(
            syncing_frame,
            text="Save",
            width=250,
            height=60,
            font=("Segoe UI", 20, "bold"),
            command=self.save_files
        ) 
        disable_button_while(self.save_button, lambda: not self.sync_done)

        self.error_log = StringRedirector()

        img = ctk.CTkImage(Image.open(resource_path("resources/err_log.png")))
        self.show_log_button = ctk.CTkButton(
            syncing_frame,
            text="",
            image=img,
            width=30,
            command=self.show_err_log
        ) 
        _hover2 = OnHover(self.show_log_button, "Show error log")

        #==================================================
        # Placing UI elements and inner containers
        #==================================================

        header_frame.pack(fill="x", anchor="n", pady=0)
        self.backup_button.pack(side=RIGHT, anchor="n", padx=5, pady=(5, 0))
        self.theme_change_button.pack(side=RIGHT, anchor="n", padx=5, pady=(5, 0))

        selection_frame.pack(**self.FRAME_0_KW)
        selection_frame.grid_columnconfigure((0, 1), weight=1)
        selection_frame.grid_rowconfigure((0, 2), weight=0)
        selection_frame.grid_rowconfigure(1, weight=1) # Only let the file display expand

        self.select_word_button.grid(row=0, column=0, sticky="new", padx=5, pady=5)
        self.select_excel_button.grid(row=0, column=1, sticky="new", padx=5, pady=5)
        self.word_file_handler.ui.grid(row=1, column=0, sticky="nsew", padx=5, pady=5)
        self.excel_file_handler.ui.grid(row=1, column=1, sticky="nsew", padx=5, pady=5)
        self.sync_button.grid(row=2, column=1, sticky="nse", padx=5, pady=(5, 10))

        syncing_frame.grid_columnconfigure(0, weight=1)
        syncing_frame.grid_columnconfigure(1, weight=0)
        syncing_frame.grid_columnconfigure(2, weight=0)
        syncing_frame.grid_rowconfigure(0, weight=1)
        syncing_frame.grid_rowconfigure(1, weight=0)
        syncing_frame.grid_rowconfigure(2, weight=0)

        self.mismatch_container.grid(row=0, column=0, columnspan=2, sticky="nsew", padx=5, pady=5)
        self.show_log_button.grid(row=1, column=0, sticky="w", padx=5, pady=(5, 0))
        self.progress_bar.grid(row=2, column=0, sticky="sew", padx=5, pady=5)
        self.save_button.grid(row=1, rowspan=2, column=1, sticky="nsew", padx=5, pady=5)
