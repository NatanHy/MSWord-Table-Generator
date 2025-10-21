import customtkinter as ctk
import sync_gui as sync
import generation_gui as gen
from gui import Tk

ASPECT_RATIO = 9 / 16
RES_X = 720
RES_Y = round(RES_X * ASPECT_RATIO)
RESOLUTION = f"{RES_X}x{RES_Y}"

class MainApp(Tk):
    def __init__(self):
        super().__init__()
        self.title("Select Operation")
        self.geometry(RESOLUTION)

        # Control buttons
        self.button_frame = ctk.CTkFrame(self)
        self.button_frame.grid_rowconfigure(0, weight=0)
        self.button_frame.grid_rowconfigure(1, weight=1)
        self.button_frame.grid_columnconfigure((0, 1), weight=1)

        self.header_label = ctk.CTkLabel(
            self.button_frame, 
            text="Select Operation",
            font=("Segoe UI", 50, "bold")
        ).grid(row=0, column=0, columnspan=2, sticky="new", padx=10, pady=10)

        ctk.CTkButton(
            self.button_frame, 
            text="Generate Tables", 
            command=self.show_gen,
            height=100,
            font=("Segoe UI", 20, "bold")
            ).grid(row=1, column=0, sticky="ew", padx=10, pady=10)
        ctk.CTkButton(
            self.button_frame, 
            text="Sync Files", 
            command=self.show_sync,
            height=100,
            font=("Segoe UI", 20, "bold")
            ).grid(row=1, column=1, sticky="ew", padx=10, pady=10)

        self.sync = sync.App(self)
        self.gen = gen.App(self)
        self.sync.run()
        self.gen.run()

        self.sync.frame_manager.callbacks[0] = self.show_buttons
        self.gen.frame_manager.callbacks[0] = self.show_buttons

        self.show_buttons()

    def show_buttons(self):
        self.title("Select Operation")
        self.button_frame.pack(expand=True, fill="both")
        self.sync.pack_forget()
        self.gen.pack_forget()

    def show_sync(self):
        self.title("File syncing")
        self.button_frame.pack_forget()
        self.gen.pack_forget()
        self.sync.pack(expand=True, fill="both")

    def show_gen(self):
        self.title("Table generation")
        self.button_frame.pack_forget()
        self.sync.pack_forget()
        self.gen.pack(expand=True, fill="both")

if __name__ == "__main__":
    ctk.set_appearance_mode("system")

    app = MainApp()
    app.mainloop()
