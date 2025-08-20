import customtkinter as ctk

class OnHover:
    def __init__(self, widget, tooltip_text="", hover_bg=None, normal_bg=None, delay=500):
        self.widget = widget
        self.tooltip_text = tooltip_text
        self.hover_bg = hover_bg
        self.normal_bg = normal_bg or widget.cget("fg_color")
        self.tooltip = None
        self.delay = delay  # ms before showing tooltip
        self.after_id = None
        
        self.widget.bind("<Enter>", self._on_enter)
        self.widget.bind("<Leave>", self._on_leave)
        self.widget.bind("<Motion>", self._on_motion)

    def _on_enter(self, event):
        if self.hover_bg:
            self.widget.configure(fg_color=self.hover_bg)
        self.after_id = self.widget.after(self.delay, lambda: self._show_tooltip(event))

    def _on_leave(self, event):
        if self.normal_bg:
            self.widget.configure(fg_color=self.normal_bg)
        if self.after_id:
            self.widget.after_cancel(self.after_id)
            self.after_id = None
        self._hide_tooltip()

    def _on_motion(self, event):
        # Move tooltip with mouse
        if self.tooltip:
            x = event.x_root + 10
            y = event.y_root + 10
            self.tooltip.geometry(f"+{x}+{y}")

    def _show_tooltip(self, event):
        if self.tooltip or not self.tooltip_text:
            return
        x = event.x_root + 10
        y = event.y_root + 10
        self.tooltip = ctk.CTkToplevel(self.widget)
        self.tooltip.wm_overrideredirect(True)  # no window decorations
        self.tooltip.geometry(f"+{x}+{y}")
        label = ctk.CTkLabel(self.tooltip, text=self.tooltip_text, fg_color="#333", text_color="white", corner_radius=5, padx=8, pady=4)
        label.pack()

    def _hide_tooltip(self):
        if self.tooltip:
            self.tooltip.destroy()
            self.tooltip = None
