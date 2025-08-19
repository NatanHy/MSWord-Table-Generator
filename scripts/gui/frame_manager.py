import customtkinter as ctk
from PIL import Image

class FrameManager:
    def __init__(
            self, 
            root, 
            frames, 
            current_frame=0, 
            frame_kwargs={}, 
            on_back_callbacks={},
            place_back_button=True,
            back_button_pos=(10, 10)
            ):
        self.frames = frames
        self.current_frame = current_frame
        self.frame_kwargs = frame_kwargs
        self.place_back_button = place_back_button
        self.callbacks = on_back_callbacks
        self.back_button_pos = back_button_pos

        # Store a reference to back image, otherwise it gets garbage collected
        self._back_img = ctk.CTkImage(light_image=Image.open("resources/back_arrow_white.png"), size=(20, 20))
        self.back_button = ctk.CTkButton(
            root, 
            image=self._back_img,
            text="",
            width=30,
            command=self.back
        )
    
    def go_to_frame(self, frame):
        from utils.gui_utils import hide_ui_element, display_ui_element #type: ignore

        if frame in self.frame_kwargs:
            kwargs = self.frame_kwargs[frame]
        else:
            kwargs = {}

        if frame < 0 or frame >= len(self.frames):
            return

        hide_ui_element(self.frames[self.current_frame])
        self.current_frame = frame
        display_ui_element(self.frames[self.current_frame], **kwargs)

        if self.current_frame != 0 and self.place_back_button:
            self.back_button.place(x=self.back_button_pos[0], y=self.back_button_pos[1])

    def back(self):
        from utils.gui_utils import hide_ui_element, display_ui_element

        if self.current_frame == 1:
            hide_ui_element(self.back_button)
            
        if self.current_frame in self.callbacks:
            self.callbacks[self.current_frame]()

        self.go_to_frame(self.current_frame - 1)