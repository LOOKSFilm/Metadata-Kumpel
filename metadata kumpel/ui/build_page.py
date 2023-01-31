import customtkinter, tkinter
from EditShareAPI import FlowMetadata
import xml.etree.ElementTree as et
import threading

import ui.mainmenu_page
from functions.build_excel import build


def buildpage(app, font, bg_color):
    def back():
        frame_main.grid_forget()
        ui.mainmenu_page.mainmenupage(app, font, bg_color)
    
    def listfields(selection):
        global xmlfile
        global is_imported
        is_imported = False
        xmlfile = ""
        listbox_fields.delete(0, "end")
        btn_import.configure(state="active")
        btn_build.configure(state="active")
        global assignfields
        assignfields = dict()
        templates = FlowMetadata.getCustomMetadataConfig()
        global fielddict
        fielddict = dict()
        for template in templates:
            if template["name"] == selection:
                fields = template["fields"]
                for i, field in enumerate(fields):
                    assignfields[field["name"]] = field["db_key"]
                    if field["name"] == "001 Identifier" or field["name"] == "014 Title Original":
                        listbox_fields.insert(i, field["name"])
                        listbox_fields.select_set(i)
                        listbox_fields.event_generate("<<ListboxSelect>>")
                        fielddict[i] = field["name"]
                        #customtkinter.CTkCheckBox(frame_map, text=field["name"], font=font).grid(sticky="NSEW")                 
                    else:
                        listbox_fields.insert(i, field["name"])
                        fielddict[i] = field["name"]
                        #customtkinter.CTkCheckBox(frame_map, text=field["name"], font=font).grid(sticky="NSEW")

    def start_build_excel():
        selects = list()
        for item in listbox_fields.curselection():
            selects.append(fielddict[item])
        threading.Thread(target=build, args=[app, selects, assignfields, is_imported, xmlfile]).start()

    def import_XML():
        global is_imported
        is_imported = True
        global xmlfile
        xmlfile = customtkinter.filedialog.askopenfilename(defaultextension=".xml", filetypes=[("XML", ".xml")])
        filename = xmlfile.split("/")[-1]
        path.set(f".../{filename}")
        tree = et.parse(xmlfile)
        root = tree.findall("clip")
        for row, clip in enumerate(root):
            row = row + 2
            metadata = clip.find("custom")
            for entry in metadata:
                try:
                    field_from_xml = entry.attrib["username"]
                except KeyError:
                    pass
                for i in fielddict.keys():
                    if fielddict[i] == field_from_xml:
                        listbox_fields.select_set(i)
                        listbox_fields.event_generate("<<ListboxSelect>>")
                    

    app.title("Metadata Kumpel: Build Excel")
    app.resizable(True,True)
    #Main Frame
    frame_main = customtkinter.CTkFrame(app, fg_color=bg_color)
    frame_main.grid_rowconfigure(0, weight=1)
    frame_main.grid_rowconfigure(1, weight=100)
    frame_main.grid_rowconfigure(2, weight=1)
    frame_main.grid_columnconfigure(0, weight=1)
    frame_main.grid(sticky="NSEW", padx=10, pady=10)
    
    #Top Bar
    frame_top = customtkinter.CTkFrame(frame_main, fg_color=bg_color)
    frame_top.grid_columnconfigure(0, weight=1)
    frame_top.grid_columnconfigure(1, weight=1)
    frame_top.grid_columnconfigure(2, weight=20)
    frame_top.grid(sticky="NSEW")
    #BTNs Top Bar
    btn_template = customtkinter.CTkLabel(frame_top, text="Select Template:", font=font, text_color="gray80")
    btn_template.grid(column=0, row=0, sticky="W")

    btn_template = customtkinter.CTkComboBox(frame_top, values=["LOOKS Archiv", "Archive Producing", "LOOKS-PROGRESS"], width=200, font=font, dropdown_font=font, variable="", command=listfields)
    btn_template.grid(column=1, row=0, sticky="W")

    btn_back = customtkinter.CTkButton(frame_top, text="Back", font=font, command=back)
    btn_back.grid(column=2, row=0, sticky="E")


    #Map Frame
    frame_map = customtkinter.CTkFrame(frame_main, fg_color="gray15", border_width=2, border_color="gray40")
    frame_map.grid(sticky="NSEW", pady=5)
    #frame_main.grid_columnconfigure(0, weight=1)
    #List Box with scrollbar
    listbox_fields= tkinter.Listbox(frame_map, background="gray15", font=font, highlightbackground="gray15", border=0, borderwidth=0, activestyle="dotbox", highlightcolor="gray15",fg="gray80", selectmode="multiple", selectbackground='#1f538d', relief="flat")
    listbox_fields.pack(side="left", expand=True, fill="both", padx=5, pady=5)

    scrollbar = customtkinter.CTkScrollbar(frame_map)
    scrollbar.pack(side="right", fill="y", padx=1, pady=5)

    scrollbar.configure(command=listbox_fields.yview)
    listbox_fields.configure(yscrollcommand=scrollbar.set)


    #Bot Frame
    frame_bot = customtkinter.CTkFrame(frame_main, height=100, fg_color=bg_color)
    frame_bot.grid_columnconfigure(0, weight=1)
    frame_bot.grid_columnconfigure(1, weight=10)
    frame_bot.grid_columnconfigure(2, weight=1)
    frame_bot.grid(sticky="EW")

    #BTNs Bot Frame
    btn_import = customtkinter.CTkButton(frame_bot, text="Import XML", font=font, state="disabled", command=import_XML)
    btn_import.grid(column=0, row=0, sticky="W")
    path = tkinter.StringVar(value="")
    label_import_path = tkinter.Label(frame_bot, textvariable=path, font=font, bg=bg_color, fg="gray80", justify="left")
    label_import_path.grid(column=1, row=0, sticky="W")
    btn_build = customtkinter.CTkButton(frame_bot, text="Build Excel", font=font, state="disabled", command=start_build_excel)
    btn_build.grid(column=2, row=0, sticky="E")

