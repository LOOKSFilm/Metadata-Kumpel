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

## Function to create EditShare search data for advanced search requests
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

def map(app, window_mapping, mapping_status, label_loading, progress_bar, frame_results, btn_cancel, excel_file, testrun, update, rename, frame_mapping_page, bg_color, VERSION, stop_event):
## Create temp log files
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

## Export log files after Mapping        
    def export_logs():
        directory = customtkinter.filedialog.askdirectory()
        for i in range(listbox_results.size()):
            filename = listbox_results.get(i)
            filepath = directory+"/"+filename+".log"
            print(directory, filepath)
            with open(filepath, "w") as f:
                f.write(databuffer[i])

## UI Back to Map Page   
    def back():
        window_mapping.grid_forget()
        ui.mapping_page.mappage(app, font=font, bg_color=bg_color, VERSION=VERSION)
    
## Get EditShares Metadatafields
    fieldsdict = FlowMetadata.getCustomMetadataFields().fields_dict

## Open Excel file 
    try:
        xlsx = openpyxl.load_workbook(excel_file, data_only=True, keep_vba=True)
    except PermissionError:
        messagebox.showerror("Permission Error", "Can not read Excel file:\n\nClose the Excel Document!\n\n")
    except NameError:
        messagebox.showerror("Permission Error", "Can not open Excel file:\n\nPlease select an Excel File!\n\n")
    sheet = xlsx.active
    rows = sheet.rows
    columns = sheet.columns

## Create dict of row 0 Metadatafields from Excel scheme: "column number": "field_000"
## Columns that don't match any EditShare metadata fields are skipped
    mappingdict = dict()
    for ir, row in enumerate(rows):
        if ir == 0:
            for ic, cell in enumerate(row):
                try:
                    mappingdict[ic] = fieldsdict[cell.value]
                except KeyError:
                    pass
    #print(json.dumps(mappingdict, indent=4))

## UI Results with Temp log files
    font = customtkinter.CTkFont(family="Hack NF", size=12, weight="bold")
    listbox_results= tkinter.Listbox(frame_results, background="gray15", font=font, borderwidth=0, selectmode="single", highlightbackground="gray15", border=0, highlightcolor="gray15", fg="gray80", selectbackground='#1f538d', relief="flat")
    listbox_results.bind("<<ListboxSelect>>", start_on_select)
    listbox_results.pack(side="left", fill="both", expand=True, pady=10, padx=5, ipadx=10)
    listbox_results.grid_columnconfigure(0, weight=1)

    sheet = xlsx.active
    rows = sheet.rows
    sheet = xlsx.active
    rows = sheet.rows
    mappings = dict()

## Mapping loop
    for ir, row in enumerate(rows):
        mapping = dict()
        ## Break point when canceling mapping
        if stop_event.is_set():
            break
## Create mappings dict; scheme: "Mapping ID": {"field_50": "data", ...}, "Mapping ID": {...}, ...
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
    #print(json.dumps(mappings, indent=4))

## Mapping 
    data = dict()
    data["combine"] = "MATCH_ANY"
    data["filters"] = list()
## When update selected search assets via 001 Identifier    
    if update == 1:
        clips = list()
        for i, mapping in enumerate(mappings):
            mapping_status.set(f"Searching Assets...")
            data["filters"].append(addSearchValue("field_50", "SEARCH_ASSETS", "EQUAL_TO", mappings[mapping]["field_50"]))
        data = json.dumps(data)
        clips = FlowMetadata.searchAdvanced(data)
## If not search Clipname   
    else:
        clips = list()
        for i, mapping in enumerate(mappings):
            mapping_status.set(f"Searching Assets...")
            data["filters"].append(addSearchValue("CLIPNAME", "SEARCH_FILES", "EQUAL_TO", mapping))
        data = json.dumps(data)
        clips = FlowMetadata.searchAdvanced(data)
    
    image = False
    if not clips:
        messagebox.showerror("Empty Search", "\nNo Asset found\n\n\n")
    id_listitem = 0
    databuffer = dict()
    skipped_assets = list()
    for i, clip in enumerate(clips):
        if stop_event.is_set():
            break
        t = PrettyTable(['Field', 'Value'])
        t.align['Field'] = "l"
        t.align['Value'] = "l"             
        if "clip_id" in clip.keys():
            metadata = FlowMetadata.getClipData(clip["clip_id"])
            try:
                if metadata["code"] == 403:
                    print(clip["clip_id"])
                    continue
            except KeyError:
                pass
            try:
                asset_id = metadata["asset"]["asset_id"]
                metadata_id = metadata["metadata"]["metadata_id"]
                capture_id = metadata["capture"]["capture_id"]
            except TypeError:
                skipped_assets.append(clip["clip_id"])
                continue
        elif "image_id" in clip.keys():
            image = True
            metadata = FlowMetadata.getImageData(clip["image_id"])
            asset_id = metadata["asset"]["asset_id"]
        else:
            continue
        if update == 1:
            try:
                id = metadata["asset"]["custom"]["field_248"]
                mapping_status.set(f"Comparing Asset:\n{id}")
            except KeyError:
                continue
        else:
            id = metadata["display_name"]
            mapping_status.set(f"Comparing Asset:\n{id}")
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
            try:
                data["custom"]["field_231"] = clip["clip_id"]
            except KeyError:
                data["custom"]["field_231"] = clip["image_id"]
            data["custom"]["field_233"] = asset_id
            data["custom"]["field_235"] = capture_id
            data["custom"]["field_62"] = ui.login_page.username
            data["custom"]["field_51"] = True
            data["custom"]["field_60"] = str(datetime.now())
            data["custom"]["field_127"] = metadata["asset"]["uuid"]
            if type(data["custom"]['field_49']) == list:
                seperator = "; "
                data["custom"]['field_49'] = seperator.join(data["custom"]['field_49'])
            try:
                if ";" in str(data["custom"]["field_129"]):
                    dates = str(data["custom"]["field_129"]).split(";")
                    newdate = str()
                    for date in dates:
                        newdate += date.strip().split(" ")[0]+"; "
                    data["custom"]["field_129"] = newdate
                elif data["custom"]["field_129"] == None:
                    pass
                else:
                    data["custom"]["field_129"] = str(data["custom"]["field_129"]).split(" ")[0]  
            except:
                pass          
            data["custom"]['field_49'] = str(data["custom"]["field_48"])
            data = json.dumps(data, indent=4)
            transmitted_json = data
            if testrun:
                data = json.loads(data)
                if not image and rename:
                    mappedData = f"Renamed Clip: {clipname}\n"
                else:
                    mappedData = "Not renamed"
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
                listbox_results.insert(id_listitem, f"{id}")
                listbox_results.itemconfig(id_listitem, foreground="yellow")
                databuffer[id_listitem] = mappedData
                id_listitem += 1
            else:
                r = FlowMetadata.updateAsset(asset_id, data)
                if r.status_code == 403:
                    messagebox.showerror("Permission Error", f"{ui.login_page.username} has no permission to write Metadata!")
                data = json.loads(data)
                reversedFielddict = dict()
                for key in fieldsdict:
                    val = fieldsdict[key]
                    reversedFielddict[val] = key
                for key in data["custom"]:
                    fieldname = reversedFielddict[key]
                    fieldval = data["custom"][key]
                    t.add_row([fieldname, fieldval])
                if not image and rename:
                    data = dict()
                    data["clip_name"] = clipname
                    data = json.dumps(data)
                    res = FlowMetadata.updateMetadata(metadata_id, data)
                    if res != "OK":
                        error = r["details"]
                        messagebox.showerror("Error 403", f"Renaming failed: {error}")
                    else:
                        pass
                    mappedData = f"Renamed Clip: {clipname}\n"
                else:
                    mappedData = "Not renamed"
                if r.status_code == 200:
                    mappedData += t.get_string(border=1)
                    mappedData += f"\nEditShare response: {r.text}\n\nTransmitted JSON\n-----------------\n"
                    mappedData += transmitted_json
                    listbox_results.insert(id_listitem, f"{id}")
                    listbox_results.itemconfig(id_listitem, foreground="green")
                    databuffer[id_listitem] = mappedData
                    id_listitem += 1
            del mappings[id]
            if len(mappings) == 0:
                break
        else:
            pass        
        
    if len(mappings) > 0:
        skipped_msg = "Assets aus der Excel wurden nicht gemappt.\nDas kann den Grund haben dass du für folgende ClipIDs keine lesberechtigung hast. Schick dieses File an Christoph, der schaut nach :)\n\n"
        skipped_msg += f"Assets nicht gemappt:\n"
        for id in mappings:
            skipped_msg += f"{id}\n"
        skipped_msg += "\nSkipped ClipIDs:\n"
        for clipID in skipped_assets:
            skipped_msg += f"{clipID}\n\n"
        listbox_results.insert(id_listitem, "Skipped Assets")
        listbox_results.itemconfig(id_listitem, foreground="red")
        databuffer[id_listitem] = skipped_msg
    mapping_status.set(f"Mapping Complete")
    progress_bar.pack_forget()
    btn_cancel.configure(text="Close", command=back)
    btn_export_logs = customtkinter.CTkButton(window_mapping, text="Export Log Files", font=font, command=export_logs)
    btn_export_logs.pack(side="right", padx=10, pady=10)
