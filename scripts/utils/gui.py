from customtkinter import ThemeManager
from PIL import ImageOps, Image
import platform, os

def open_folder(folder_path):
    if platform.system() == "Windows":
        os.startfile(folder_path)
    elif platform.system() == "Darwin":  # macOS
        os.system(f"open '{folder_path}'")
    else:  # Linux
        os.system(f"xdg-open '{folder_path}'")

def color_filter(img, clr):
    # Split channels to recover alpha later
    img = img.convert("RGBA")
    r, g, b, alpha = img.split()

    # Convert to grayscale for ImageOps
    grayscale = Image.merge("RGB", (r, g, b)).convert("L")
    colored = ImageOps.colorize(grayscale, black=clr[1], white=clr[0])

    # Add alpha channel back
    colored.putalpha(alpha)

    return colored

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