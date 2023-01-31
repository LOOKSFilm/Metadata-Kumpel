import customtkinter
import tkinter

class LoadingAnimation(customtkinter.CTkFrame):
    def __init__(self, parent, frames, delay, font, bg_color):
        customtkinter.CTkFrame.__init__(self, parent)
        self.frames = frames
        self.delay = delay
        self.current_frame = 0
        self.label = customtkinter.CTkLabel(self, text="", font=font, bg_color=bg_color)
        self.label.pack()
        self.update_frame()

    def update_frame(self):
        self.current_frame = (self.current_frame + 1) % len(self.frames)
        self.label.configure(text=self.frames[self.current_frame])
        self.after(self.delay, self.update_frame)

class ToggleText(customtkinter.CTkFrame):
    def __init__(self, parent, buttontext, text, font):
        super().__init__(parent)
        self.istext = False
        self.textvar = tkinter.StringVar(value=text)
        self.buttontext = buttontext
        self.font = font
        self.create_widgets()

    def create_widgets(self):
        self.button = customtkinter.CTkButton(self, text=self.buttontext, command=self.toggle_text, font=self.font, bg_color="gray30", fg_color="gray30")
        self.button.pack(expand=True, fill="both")

    def toggle_text(self):
        if not self.istext:
            self.text = customtkinter.CTkLabel(self, height=5, width=30, textvariable=self.textvar, font=self.font)
            self.istext = True
            self.text.pack(expand=True, fill="both")
        else:
            self.text.pack_forget()
            self.istext = False