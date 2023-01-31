import customtkinter
import tkinter
from EditShareAPI import FlowMetadata
import threading

import ui.mainmenu_page
from functions.map_excel import map

def mappage(app, font, bg_color):
    global btn_testrun
    global btn_start_mapping
    def back():
        frame_mapping_page.grid_forget()
        ui.mainmenu_page.mainmenupage(app, font, bg_color)

    def excel_select():
        global excel_file
        excel_file = customtkinter.filedialog.askopenfilename(defaultextension=".xlsm", filetypes=[("Excel", ".xlsm")])
        if excel_file == "":
            excel_filename.set("No Excel selected")
        else:
            excel_filename.set(".../"+str(excel_file.split("/")[-1]))
        on_select("")
        options_select(options_collect._current_value)

    def on_select(event):
        if excel_filename.get() != "No Excel selected" and len(listbox_options.curselection()) != 0:
            btn_start_mapping.configure(state="active")
            btn_testrun.configure(state="active")
        else:
            btn_start_mapping.configure(state="disabled")
            btn_testrun.configure(state="disabled")

    def options_select(value):
        btn_start_mapping.configure(state="disabled")
        btn_testrun.configure(state="disabled")
        if value == "Collect Assets via Mediaspace" or  value == "Collect Assets via Source":
            if input_prefix.winfo_exists():
                input_prefix.grid_forget()
            frame_option_settings.grid(column=0, row=2, sticky="NSEW")
            listbox_options.delete(0,"end")
            if value == "Collect Assets via Mediaspace":
                for i, mediaspace in enumerate(FlowMetadata.getMediaSpaces()):
                    listbox_options.insert(i, mediaspace["name"])
                    #print(mediaspace["name"])
            if value == "Collect Assets via Source":
                fields = FlowMetadata.getCustomMetadataFields()
                for field in fields:
                    if field["name"][0:4] == "006 ":
                        for i, value in enumerate(field["allowed_values"]["values"]):
                            listbox_options.insert(i, value["value"])
        else:
            frame_option_settings.grid_forget()
            #frame_option_settings.configure(border_color=bg_color)
            if value == "Collect Assets by Prefix":
                input_prefix.grid(column=0, row=2, sticky="NW")
                if excel_filename.get() != "No Excel selected":
                    btn_start_mapping.configure(state="active")
                    btn_testrun.configure(state="active")
            else:
                if excel_filename.get() != "No Excel selected":
                    btn_start_mapping.configure(state="active")
                    btn_testrun.configure(state="active")
                if input_prefix.winfo_exists():
                    input_prefix.grid_forget()
    def start_testrun():
        start_mapping(True)

    def start_mapping(testrun = False):
        mapping_option = options_collect._current_value
        btn_start_mapping.configure(state="disabled")
        btn_testrun.configure(state="disabled")
        if mapping_option == "Collect Assets via Mediaspace" or options_collect._current_value == "Collect Assets via Source":
            values = list()
            for i in listbox_options.curselection():
                values.append(listbox_options.get(i))
        elif mapping_option == "Collect Assets by Prefix":
            values = input_prefix.get()
        else:
            values = None
        update =toggle_update.get()
        threading.Thread(target=map, args=[app, mapping_option, values, excel_file, testrun, update]).start()




    app.title("Metadata Kumpel: Map Excel")
    app.resizable(True,True)
    # app.grid_columnconfigure(0, weight=1)
    # app.grid_rowconfigure(0, weight=1)
    #Mainframe
    frame_mapping_page = customtkinter.CTkFrame(app, fg_color=bg_color)
    frame_mapping_page.rowconfigure(0, weight=1)
    frame_mapping_page.rowconfigure(1, weight=1)
    frame_mapping_page.rowconfigure(2, weight=100)
    frame_mapping_page.rowconfigure(3, weight=1)
    frame_mapping_page.rowconfigure(4, weight=1)
    #frame_mapping_page.rowconfigure(3, weight=1)
    frame_mapping_page.columnconfigure(0, weight=1)
    frame_mapping_page.grid(sticky="NSEW", padx=10, pady=10)
    

    #Bar top
    frame_top = customtkinter.CTkFrame(frame_mapping_page, fg_color=bg_color)
    frame_top.columnconfigure(0, weight=1)
    frame_top.columnconfigure(1, weight=10)
    frame_top.columnconfigure(2, weight=1)
    frame_top.grid(sticky="EW", row=0)
    btn_select_excel = customtkinter.CTkButton(frame_top, text="Select Excel", command=excel_select, font=font)
    btn_select_excel.grid(column=0, row=0, sticky="W")
    excel_filename = tkinter.StringVar(value="No Excel selected")
    label_path_excel = tkinter.Label(frame_top, textvariable=excel_filename, font=font, bg=bg_color, fg="gray80")
    label_path_excel.grid(column=1, row=0, sticky="W")
    btn_back = customtkinter.CTkButton(frame_top, text="Back", command=back, font=font)
    btn_back.grid(column=2, row=0, sticky="E")

    options = ["Collect Assets via Mediaspace", "Collect Assets via Source", "Collect Assets by Prefix", "Search Asset by ID (Slow)"]
    options_collect = customtkinter.CTkOptionMenu(frame_mapping_page, font=font, values=options, command=options_select, dropdown_font=font)
    options_collect.grid(column=0, row=1, pady=10, sticky="W")

    #Options Frame
    frame_option_settings = customtkinter.CTkFrame(frame_mapping_page, fg_color="gray15", border_width=2, border_color="gray40")
    frame_option_settings.grid(column=0, row=2, sticky="NSEW")
    listbox_options= tkinter.Listbox(frame_option_settings, background="gray15", font=font, highlightbackground="gray15", border=0, borderwidth=0, activestyle="dotbox", highlightcolor="gray15",fg="gray80", selectmode="single", selectbackground='#1f538d', relief="flat")
    listbox_options.bind("<<ListboxSelect>>", on_select)
    scrollbar = customtkinter.CTkScrollbar(frame_option_settings)
    scrollbar.configure(command=listbox_options.yview)
    listbox_options.configure(yscrollcommand=scrollbar.set)
    listbox_options.pack(side="left", expand=True, fill="both", padx=5, pady=5)
    scrollbar.pack(side="right", fill="y", padx=1, pady=5)
    input_prefix = customtkinter.CTkEntry(frame_mapping_page, placeholder_text="Insert Prefix", font=font, width=500)
    

    #Update Toggle
    toggle_update = customtkinter.CTkSwitch(frame_mapping_page, font=font, text="Update Metadata")
    toggle_update.grid(row=3)

    #Bar Bot
    frame_bot = customtkinter.CTkFrame(frame_mapping_page, fg_color=bg_color)
    frame_bot.columnconfigure(0, weight=1)
    frame_bot.columnconfigure(1, weight=10)
    frame_bot.columnconfigure(2, weight=1)
    frame_bot.grid(sticky="EW", row=4)
    
    btn_testrun = customtkinter.CTkButton(frame_bot, text="Testrun", command=start_testrun, font=font, state="disabled")
    btn_testrun.grid(column=0, row=0, sticky="W")
    btn_start_mapping = customtkinter.CTkButton(frame_bot, text="Start Mapping", command=start_mapping, font=font, state="disabled")
    btn_start_mapping.grid(column=2, row=0, sticky="E")

    options_select("Collect Assets via Mediaspace")