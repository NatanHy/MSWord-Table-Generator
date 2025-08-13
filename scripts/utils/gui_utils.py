from customtkinter import ThemeManager, CTkButton, CTkBaseClass, LEFT, RIGHT, TOP, BOTTOM
from PIL import ImageOps, Image
import platform, os
import customtkinter as ctk
from typing import Callable, List
from gui import SelectedFilesHandler, PopUpWindow

def wrong_files_popup(root, err_message):
    def f(file_handler : 'SelectedFilesHandler', file_paths : List[str]):
        popup_win = PopUpWindow(root, "Wrong file type", err_message)
        return _show_wrong_file_popup(popup_win, file_handler, file_paths)
    return f

def _show_wrong_file_popup(popup_win : 'PopUpWindow', file_handler : 'SelectedFilesHandler', file_paths : List[str]):
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

def disable_button_while(button : CTkButton, condition : Callable[[], bool], poll_rate_ms=100):
    # If condition is true, keep button disabled
    if condition():
        disable_button(button)
        button.after(poll_rate_ms, lambda: disable_button_while(button, condition)) # Poll again 100ms later
        return
    
    # Otherwise enable button
    enable_button(button)
    button.after(poll_rate_ms, lambda: disable_button_while(button, condition)) # Poll again 100ms later

def open_folder(folder_path):
    if platform.system() == "Windows":
        os.startfile(folder_path)
    elif platform.system() == "Darwin":  # macOS
        os.system(f"open '{folder_path}'")
    else:  # Linux
        os.system(f"xdg-open '{folder_path}'")

def color_filter(img : Image.Image, clr):
    # Split channels to recover alpha later
    img = img.convert("RGBA")
    r, g, b, alpha = img.split()

    # Convert to grayscale for ImageOps
    grayscale = Image.merge("RGB", (r, g, b)).convert("L")
    colored = ImageOps.colorize(grayscale, black=clr[1], white=clr[0])

    # Add alpha channel back
    colored.putalpha(alpha)

    return colored

def disable_button(button : CTkButton): 
    if button._state != "disabled":
        button.configure(state="disabled", fg_color="#5a5a5a", text_color="gray80")

def enable_button(button : CTkButton):
    if button._state != "normal":
        clr = ThemeManager.theme["CTkButton"]["fg_color"]
        button.configure(state="normal", fg_color=clr, text_color="white")

def display_ui_element(elm : CTkBaseClass, **kwargs):
    elm.pack(**kwargs)

def hide_ui_element(elm):
    elm.pack_forget()
    elm.place_forget()
    elm.grid_forget()