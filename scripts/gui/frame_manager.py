import customtkinter as ctk
from PIL import Image

from utils.files import resource_path

class FrameManager:
    def __init__(
            self, 
            root, 
            frames, 
            current_frame=0, 
            frame_kwargs={}, 
            on_back_callbacks={},
            back_button_pos=(10, 10)
            ):
        self.frames = frames
        self.current_frame = current_frame
        self.frame_kwargs = frame_kwargs
        self.callbacks = on_back_callbacks
        self.back_button_pos = back_button_pos

        # Store a reference to back image, otherwise it gets garbage collected
        self._back_img = ctk.CTkImage(light_image=Image.open(resource_path("resources/back_arrow_white.png")), size=(20, 20))
        self.back_button = ctk.CTkButton(
            root, 
            image=self._back_img,
            text="",
            width=30,
            command=self.back
        )

        self.back_button.place(x=self.back_button_pos[0], y=self.back_button_pos[1])
    
    def go_to_frame(self, frame):
        from utils.gui_utils import hide_ui_element, display_ui_element

        if frame < 0 or frame >= len(self.frames):
            return
        
        if frame in self.frame_kwargs:
            kwargs = self.frame_kwargs[frame]
        else:
            kwargs = {}

        hide_ui_element(self.frames[self.current_frame])
        self.current_frame = frame
        display_ui_element(self.frames[self.current_frame], **kwargs)

        self.back_button.place(x=self.back_button_pos[0], y=self.back_button_pos[1])

    def back(self):
        if self.current_frame in self.callbacks:
            self.callbacks[self.current_frame]()

        self.go_to_frame(self.current_frame - 1)