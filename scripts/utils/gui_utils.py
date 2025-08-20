import os
import platform
from typing import Callable, List, TYPE_CHECKING

import customtkinter as ctk
from customtkinter import ThemeManager, CTkButton, CTkBaseClass, LEFT, RIGHT
from PIL import ImageOps, Image

if TYPE_CHECKING:
    from gui import SelectedFilesHandler, PopUpWindow

def switch_theme():
    current = ctk.get_appearance_mode()
    if current == "Dark":
        ctk.set_appearance_mode("Light")
    else:
        ctk.set_appearance_mode("Dark")

def wrong_files_popup(root, err_message):
    def f(file_handler : 'SelectedFilesHandler', file_paths : List[str]):
        from gui import PopUpWindow
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

def blend_colors(fg_hex, bg_hex, alpha):
    """Blend fg_hex over bg_hex with given alpha (0.0-1.0)."""
    fg = tuple(int(fg_hex[i:i+2], 16) for i in (1, 3, 5))
    bg = tuple(int(bg_hex[i:i+2], 16) for i in (1, 3, 5))
    blended = tuple(int((alpha * fg_c) + ((1 - alpha) * bg_c)) for fg_c, bg_c in zip(fg, bg))
    return "#{:02x}{:02x}{:02x}".format(*blended)

def _to_hex(widget, color: str) -> str:
    r, g, b = widget.winfo_rgb(color)
    return f"#{r//256:02x}{g//256:02x}{b//256:02x}"

def _normalize_color(widget, color):
    if isinstance(color, list):
        return [_to_hex(widget, c) for c in color]
    return _to_hex(widget, color)

def get_color(widget, widget_name, field_name):
    match ctk.get_appearance_mode():
        case "Light":
            return _normalize_color(widget, ThemeManager.theme[widget_name][field_name])[0]
        case "Dark":
            return _normalize_color(widget, ThemeManager.theme[widget_name][field_name])[1]

def disable_button(button : CTkButton): 
    if button._state != "disabled":
        button.configure(
            state="disabled",
            fg_color=("gray80", "#5a5a5a"),   # light mode, dark mode
            text_color=("gray40", "gray80")   # light mode, dark mode
        )

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