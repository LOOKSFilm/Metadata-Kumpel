from EditShareAPI import FlowMetadata
import openpyxl
import re
import customtkinter
import tkinter
from tkinter import messagebox
import json
from datetime import datetime
import tempfile
import subprocess
import threading
from prettytable import PrettyTable

import ui.login_page
import ui.mapping_page

def addSearchValue(fieldname, group, match, search):
    field = dict()
    field["field"] = dict()
    if fieldname.startswith("field_"):
        field["custom_field"] = fieldname
        field["field"]["fixed_field"] = f"CUSTOM_{fieldname}"
        field["field"]["group"] = group
        field["match"] = match
        field["search"] = search
    else:
        field["field"]["fixed_field"] = fieldname
        field["field"]["group"] = group
        field["match"] = match
        field["search"] = search
    return field

def map(app, mapping_option, selectvalues, excel_file, testrun, update):

    def on_select():
        mapping_result_name = listbox_results.get(listbox_results.curselection())+"__"
        with tempfile.NamedTemporaryFile(mode='w+', delete=False, prefix=mapping_result_name) as temp:
            # Write the string to the temporary file
            temp.write(databuffer[listbox_results.curselection()[0]])
            temp.seek(0)
            print(temp.read())
            # Do something with the temporary file, for example, pass it to an editor
            subprocess.run(["notepad.exe", temp.name])
    def start_on_select(event):
        threading.Thread(target=on_select).start()
        
    def export_logs():
        directory = customtkinter.filedialog.askdirectory()
        for i in range(listbox_results.size()):
            filename = listbox_results.get(i)
            filepath = directory+"/"+filename+".log"
            print(directory, filepath)
            with open(filepath, "w") as f:
                f.write(databuffer[i])

    fields = FlowMetadata.getCustomMetadataFields()
    try:
        xlsx = openpyxl.load_workbook(excel_file, data_only=True, keep_vba=True)
    except PermissionError:
        messagebox.showerror("Permission Error", "Can not read Excel file:\n\nClose the Excel Document!\n\n")
    except NameError:
        messagebox.showerror("Permission Error", "Can not open Excel file:\n\nPlease select an Excel File!\n\n")
    sheet = xlsx.active
    rows = sheet.rows
    columns = sheet.columns

    fieldsdict = dict()
    for field in fields:
        if re.match("[0-9][0-9][0-9]", field["name"][:3]):
            fieldsdict[field["name"]] = field["db_key"]
    mappingdict = dict()
    for ir, row in enumerate(rows):
        if ir == 0:
            for ic, cell in enumerate(row):
                try:
                    mappingdict[ic] = fieldsdict[cell.value]
                except KeyError:
                    pass

    sheet = xlsx.active
    rows = sheet.rows
    sheet = xlsx.active
    rows = sheet.rows
    mappings = dict()

    size, x_coord, y_coord = app.geometry().split("+")

    font = customtkinter.CTkFont(family="Hack NF", size=12, weight="bold")
    window_mapping = customtkinter.CTkToplevel()
    window_mapping.title("Mapping Excel")
    window_mapping.geometry(f"{400}x{70}+{str(int(x_coord)+100)}+{str(int(y_coord)+300)}")
    window_mapping.overrideredirect(True)
    window_mapping.attributes("-topmost", True)
    mapping_status = tkinter.StringVar(value="Reading Excel")
    label_loading = customtkinter.CTkLabel(window_mapping, textvariable=mapping_status , font=font)
    label_loading.pack()
    progress_bar = customtkinter.CTkProgressBar(window_mapping, mode="indeterminte", width=350)
    progress_bar.pack()
    progress_bar.start()

    for ir, row in enumerate(rows):
        mapping = dict()
        for ic, cell in enumerate(row):
            if ir > 0:
                if ic == 0:
                    if cell.value == None:
                        break
                    mapping_id = str(cell.value)
                    mappings[mapping_id] = dict()
                try:
                    field = FlowMetadata.getCustomMetadataField(mappingdict[ic])
                    if field["multi_select"]:
                        if type(cell.value) == str:
                            values = cell.value.split(";")
                            listvalue = list()
                            for value in values:
                                listvalue.append(value.strip()) 
                            mappings[mapping_id][mappingdict[ic]] = listvalue
                        else:
                           mappings[mapping_id][mappingdict[ic]] = None
                    else:
                        if type(cell.value) == str:
                            cell.value = cell.value.strip()
                            #print(cell.value)
                        mappings[mapping_id][mappingdict[ic]] = cell.value
                except KeyError:
                    pass
    data = dict()
    data["combine"] = "MATCH_ALL"
    data["filters"] = list()
    #dpg.add_input_text(default_value="Searching Assets: ", parent="mapMsg", tag="searchingAsset")
    
    if mapping_option == "Collect Assets by Prefix":
        prefix = selectvalues
        mapping_status.set(f"Searching Assets with Prefix: {prefix}")
        data["filters"].append(addSearchValue("CLIPNAME", "SEARCH_FILES", "BEGINS_WITH", prefix))
        data = json.dumps(data)
        clips = FlowMetadata.searchAdvanced(data)
            
    elif mapping_option == "Collect Assets via Mediaspace":
        mediaspaces = list()
        for value in selectvalues:
            mediaspaces.append(value)

        clips = list()
        for mediaspace in mediaspaces:
            mapping_status.set(f"Searching for Assets on Mediaspace: {mediaspace}")
            mediaspaceclips = FlowMetadata.getMediaSpaceClips(mediaspace)
            clips += mediaspaceclips
          

    elif mapping_option == "Collect Assets via Source":
        clips = list()
        for value in selectvalues:
            mapping_status.set(f"Searching for Assets with Source: {value}")
            data = dict()
            data["combine"] = "MATCH_ALL"
            data["filters"] = list()
            print("searching "+value)
            data["filters"].append(addSearchValue("field_55", "SEARCH_ASSETS", "EQUAL_TO", value))
            data = json.dumps(data)
            clips += FlowMetadata.searchAdvanced(data)
    
    else:
        clips = list()
        for i, mapping in enumerate(mappings):
            mapping_status.set(f"Searching Asset with Clipname: {mapping}")
            data = dict()
            data["combine"] = "MATCH_ALL"
            data["filters"] = list()
            data["filters"].append(addSearchValue("CLIPNAME", "SEARCH_FILES", "EQUAL_TO", mapping))
            data = json.dumps(data)
            clips += FlowMetadata.searchAdvanced(data)

    window_mapping.geometry(f"{800}x{600}")
    window_mapping.overrideredirect(False)
    window_mapping.attributes("-topmost", False)
    frame_results = customtkinter.CTkFrame(window_mapping, fg_color="gray15", border_width=2, border_color="gray40")
    frame_results.pack(fill="both", expand=True, padx=10, pady=10)
    listbox_results= tkinter.Listbox(frame_results, background="gray15", font=font, borderwidth=0, selectmode="single", highlightbackground="gray15", border=0, highlightcolor="gray15", fg="gray80", selectbackground='#1f538d', relief="flat")
    listbox_results.bind("<<ListboxSelect>>", start_on_select)
    listbox_results.pack(side="left", fill="both", expand=True, pady=10, padx=5, ipadx=10)
    listbox_results.grid_columnconfigure(0, weight=1)
    scrollbar = customtkinter.CTkScrollbar(frame_results)
    scrollbar.pack(side="right", fill="y", padx=1, pady=10)

    

    if not clips:
        messagebox.showerror("Empty Search", "\nNo Asset found\n\n\n")
    id_listitem = 0
    databuffer = dict()
    for i, clip in enumerate(clips):
        t = PrettyTable(['Field', 'Value'])
        t.align['Field'] = "l"
        t.align['Value'] = "l"             
        if "clip_id" in clip.keys():
            metadata = FlowMetadata.getClipData(clip["clip_id"])
            asset_id = metadata.asset["asset_id"]
            metadata_id = metadata.metadata["metadata_id"]
            capture_id = metadata.capture["capture_id"]
            if update == 1:
                try:
                    id = metadata.asset["custom"]["field_248"]
                    mapping_status.set(f"Updating Asset-Metadata: {id}")
                except KeyError:
                    continue
            else:
                id = metadata.display_name
                mapping_status.set(f"Mapping Asset-Metadata: {id}")
            if id in mappings:
                try:
                    identifier = mappings[id]["field_50"] 
                except KeyError:
                    messagebox.showerror("Key Error ID", f'Excel Error:\n\n{id}: "001 Identifier" missing\n\n')
                try:
                    name = mappings[id]["field_63"].replace(":", "").replace("\\", "-").replace("/", "-").replace(";","").replace(",", "").replace(".", "").replace("?","").replace("!","").replace("ß","sz").replace("ä","ae").replace("ü","ue").replace("ö","oe").replace("'","").replace('"',"")
                except KeyError:
                    messagebox.showerror("Key Error Title", f'Excel Error:\n\n{id}: "014 Title Original" missing\n\n')
                clipname = f"{identifier}__{name}"
                data = dict()
                data["custom"] = mappings[id]
                data["custom"]["field_248"] = id
                data["custom"]["field_231"] = clip["clip_id"]
                data["custom"]["field_233"] = asset_id
                data["custom"]["field_235"] = capture_id
                data["custom"]["field_62"] = ui.login_page.username
                data["custom"]["field_51"] = True
                data["custom"]["field_60"] = str(datetime.now())
                data["custom"]["field_127"] = metadata.asset["uuid"]
                data = json.dumps(data, indent=4)
                transmitted_json = data
                if testrun:
                    data = json.loads(data)
                    mappedData = f"Renamed Clip: {clipname}\n"
                    reversedFielddict = dict()
                    for key in fieldsdict:
                        val = fieldsdict[key]
                        reversedFielddict[val] = key
                    for key in data["custom"]:
                        fieldname = reversedFielddict[key]
                        fieldval = data["custom"][key]
                        t.add_row([fieldname, fieldval])
                    mappedData += t.get_string(border=1)
                    mappedData += f"\n\nTransmitted JSON\n-----------------\n"
                    mappedData += transmitted_json
                    listbox_results.configure(fg="yellow")
                    listbox_results.insert(id_listitem, f"{id}")
                    databuffer[id_listitem] = mappedData
                    id_listitem += 1

                else:
                    r = FlowMetadata.updateAsset(asset_id, data)
                    if r.status_code == 403:
                        messagebox.showerror("Permission Error", f"{ui.login_page.username} has no permission to write Metadata!")
                    data = json.loads(data)
                    mappedData = f"Renamed Clip: {clipname}\n"
                    reversedFielddict = dict()
                    for key in fieldsdict:
                        val = fieldsdict[key]
                        reversedFielddict[val] = key
                    for key in data["custom"]:
                        fieldname = reversedFielddict[key]
                        fieldval = data["custom"][key]
                        t.add_row([fieldname, fieldval])
                    mappedData += t.get_string(border=1)
                    mappedData += f"\nEditShare response: {r.text}\n\nTransmitted JSON\n-----------------\n"
                    mappedData += transmitted_json
                    
                    data = dict()
                    data["clip_name"] = clipname
                    data = json.dumps(data)
                    r = FlowMetadata.updateMetadata(metadata_id, data)
                    if r == "OK":
                        listbox_results.configure(fg="green")
                        listbox_results.insert(id_listitem, f"{id}")
                        databuffer[id_listitem] = mappedData
                        id_listitem += 1
                    errorcode = 200
                    try:
                        errorcode = r["code"]
                    except:
                        pass
                    if errorcode == 403:
                        error = r["details"]
                        messagebox.showerror("Error 403", f"Renaming failed: {error}")
            else:
                pass
    mapping_status.set(f"Mapping Complete")
    progress_bar.pack_forget()
    ui.mapping_page.btn_start_mapping.configure(state="active")
    ui.mapping_page.btn_testrun.configure(state="active")
    btn_close = customtkinter.CTkButton(window_mapping, text="Close", font=font, command=lambda:window_mapping.destroy())
    btn_export_logs = customtkinter.CTkButton(window_mapping, text="Export Log Files", font=font, command=export_logs)
    btn_close.pack(side="left", padx=10, pady=10)
    btn_export_logs.pack(side="right", padx=10, pady=10)
