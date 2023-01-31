import customtkinter
from ui.login_page import loginpage 


customtkinter.set_default_color_theme("dark-blue")
app = customtkinter.CTk()
app.configure(bg_color="gray10", fg_color="gray10")
app.grid_columnconfigure(0, weight=1)
app.grid_rowconfigure(0, weight=1)
app.geometry(f"{400}x{275}+{500}+{300}")
bg_color = "gray10"
VERSION = "1.0"
font = customtkinter.CTkFont(family="Hack NF", size=12, weight="bold")
#font = customtkinter.CTkFont(family="MesloLGS NFM Standart", size=12, weight="bold")
loginpage(app, font, bg_color)
app.mainloop()
