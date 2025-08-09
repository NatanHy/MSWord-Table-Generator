from customtkinter import ThemeManager

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