import customtkinter
from PIL import Image
from ui.build_page import buildpage
from ui.mapping_page import mappage

import os
import sys

#VERSION = "v1.5.23"

def resource_path(relative_path):
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)

def mainmenupage(app, font, bg_color, VERSION):
    def goto_buildpage():
        frame_main.grid_forget()
        buildpage(app, font, bg_color, VERSION)
    def goto_mappage():
        frame_main.grid_forget()
        mappage(app, font, bg_color, VERSION)

    app.geometry(f"{800}x{600}")
    app.resizable(False,False)
    app.title("Metadata Kumpel: Mainmenu")

    frame_main = customtkinter.CTkFrame(app, fg_color=bg_color)
    frame_main.grid_rowconfigure(0, weight=1)
    frame_main.grid_rowconfigure(1, weight=2)
    frame_main.grid_columnconfigure(0, weight=1)
    frame_main.grid(sticky="NSEW", row=0, column=0)

    frame_top = customtkinter.CTkFrame(frame_main, fg_color=bg_color)
    frame_top.grid(row=0, column=0, sticky="NSEW")

    frame_bot = customtkinter.CTkFrame(frame_main, fg_color=bg_color)
    frame_bot.grid_columnconfigure(0, weight=1)
    frame_bot.grid_columnconfigure(1, weight=1)
    frame_bot.grid(row=1,column=0,sticky="NSEW")

    image_width = 800
    image = customtkinter.CTkImage(dark_image=Image.open(resource_path("images/metadata_kumpel.png")), size=(image_width,image_width/3.3))
    label = customtkinter.CTkLabel(frame_top,image=image,text="")
    label.grid(column=0, sticky="NSEW")

    btn_build = customtkinter.CTkButton(frame_bot, text="Build Excel", font=font, command=goto_buildpage)
    btn_map = customtkinter.CTkButton(frame_bot, text="Map Metadata", font=font, command=goto_mappage)
    btn_build.grid(column=0, row=0, sticky="NSEW", ipady=80, padx=20)
    btn_map.grid(column=1, row=0,  sticky="NSEW", ipady=80, padx=20)
    label_version = customtkinter.CTkLabel(frame_main, text=VERSION, font=font, text_color="gray80")
    label_version.grid(column=0, row=2, padx=5, pady=5, sticky="E")
    