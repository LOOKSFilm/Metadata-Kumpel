import customtkinter
import tkinter
from EditShareAPI import FlowMetadata
import threading

import ui.mainmenu_page
from functions.map_excel import map

def mappage(app, font, bg_color, VERSION):
    global btn_start_mapping
    def back():
        frame_mapping_page.grid_forget()
        ui.mainmenu_page.mainmenupage(app, font, bg_color, VERSION)

    def excel_select():
        global excel_file
        excel_file = customtkinter.filedialog.askopenfilename(defaultextension=".xlsm", filetypes=[("Excel", ".xlsm")])
        if excel_file == "":
            btn_select_excel.configure(text="Select Excel")
        else:
            btn_select_excel.configure(text=".../"+str(excel_file.split("/")[-1]))
        on_select("")

    def on_select(event):
        if excel_file != "":
            btn_start_mapping.configure(state="active")
        else:
            btn_start_mapping.configure(state="disabled")

    def start_mapping():
        rename = toggle_rename.get()
        update = toggle_update.get()
        testrun = False
        if toggle_testrun.get() == 1:
            testrun = True
        frame_mapping_page.grid_forget()
        window_mapping = customtkinter.CTkFrame(app, fg_color=bg_color, bg_color=bg_color, border_color=bg_color)
        window_mapping.grid(sticky="NSEW", padx=10, pady=10)
        mapping_status = tkinter.StringVar(value="Reading Excel...")
        label_loading = customtkinter.CTkLabel(window_mapping, textvariable=mapping_status , font=font)
        label_loading.pack()
        progress_bar = customtkinter.CTkProgressBar(window_mapping, mode="indeterminte", width=350)
        progress_bar.pack()
        progress_bar.start()
        frame_results = customtkinter.CTkFrame(window_mapping, fg_color="gray15", border_width=2, border_color="gray40")
        frame_results.pack(fill="both", expand=True, padx=10, pady=10)
        scrollbar = customtkinter.CTkScrollbar(frame_results)
        scrollbar.pack(side="right", fill="y", padx=2, pady=10)
        stop_event = threading.Event()
        def cancel():
            stop_event.set()
            btn_cancel.configure(text="Canceling Mapping...")
        btn_cancel = customtkinter.CTkButton(window_mapping, text="Cancel", font=font, command=cancel)
        btn_cancel.pack(side="left", padx=10, pady=10)
        map_thread = threading.Thread(target=map, args=[app, window_mapping, mapping_status, label_loading, progress_bar, frame_results, btn_cancel, excel_file, testrun, update, rename, frame_mapping_page, bg_color, VERSION, stop_event])
        map_thread.start()
        




    app.title("Metadata Kumpel: Map Excel")
    app.resizable(True,True)

    #Mainframe
    frame_mapping_page = customtkinter.CTkFrame(app, fg_color=bg_color, bg_color=bg_color, border_color=bg_color)
    frame_mapping_page.rowconfigure(0, weight=1)
    frame_mapping_page.rowconfigure(1, weight=10)
    frame_mapping_page.rowconfigure(2, weight=20)
    frame_mapping_page.rowconfigure(3, weight=70)
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
    btn_back = customtkinter.CTkButton(frame_top, text="Back", command=back, font=font)
    btn_back.grid(column=2, row=0, sticky="E")

    btn_select_excel = customtkinter.CTkButton(frame_mapping_page, text="Select Excel", command=excel_select, font=font)
    btn_select_excel.grid(column=0, row=1, sticky="WS")
    #Options Frame
    
    frame_option_settings = customtkinter.CTkFrame(frame_mapping_page, border_width=1, border_color="gray80")
    frame_option_settings.grid(column=0, row=2, sticky="NSEW", pady=20)

    opt_label = customtkinter.CTkLabel(frame_option_settings, text="Options", font=font)
    opt_label.pack(anchor="w", padx=5, pady=2)
    #Rename Toggle
    toggle_rename = customtkinter.CTkSwitch(frame_option_settings, font=font, text="Rename", text_color="gray80")
    toggle_rename.pack(anchor="w", padx=10, pady=5)
    toggle_rename.select()
    #Update Toggle
    toggle_update = customtkinter.CTkSwitch(frame_option_settings, font=font, text="Update Metadata", text_color="gray80")
    toggle_update.pack(anchor="w", padx=10, pady=5)
    #Testrun Toggle
    toggle_testrun = customtkinter.CTkSwitch(frame_option_settings, font=font, text="Testrun", text_color="gray80")
    toggle_testrun.pack(anchor="w", padx=10, pady=5)
    toggle_testrun.select()

    #Bar Bot
    frame_bot = customtkinter.CTkFrame(frame_mapping_page, fg_color=bg_color)
    frame_bot.columnconfigure(0, weight=1)
    frame_bot.columnconfigure(1, weight=10)
    frame_bot.columnconfigure(2, weight=1)
    frame_bot.grid(sticky="EW", row=4)
    
    btn_start_mapping = customtkinter.CTkButton(frame_bot, text="Start Mapping", command=start_mapping, font=font, state="disabled")
    btn_start_mapping.grid(columnspan=3, row=0, sticky="ew")

