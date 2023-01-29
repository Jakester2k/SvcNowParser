import sys
import tkinter as tk
import tkinter.ttk as ttk
from tkinter.constants import *
import contextlib
import csv
import re
from tkinter import *
import os
from shutil import move
import pyperclip
import ParserGUI
import webbrowser
from collections import OrderedDict

def file_handling(source):

    destination = os.getcwd()

    for file in os.listdir(source):
        if file in [
            "wm_task.csv",
            "alm_transfer_order_line.csv",
        ] and os.path.exists(os.path.join(source, file)):
            move(os.path.join(source, file), os.path.join(destination, file))

    # Open wm_task.csv and extract the individual records into Calls.
    with open("wm_task.csv", "r") as file:
        reader = csv.DictReader(file)
        Calls = [OrderedDict(row) for row in reader]

    Calls = data_cleanup(Calls)

    # Create a list of the Titles
    Titles = ["Status", "Number", "DSPNum", "Name", "Comp", "Addr1", "Addr2", "City", "State", "Zip", "Phone", "Phone2", "Email", "AltContact", "AltPhone", "AltPhone2", "AltEmail", "Serial", "Model", "Issue", "Notes"]
    # Create a dictionary that maps the original field names to the Titles
    field_name_mapping = {field_name: Titles[i] for i, field_name in enumerate(Calls[0].keys())}

    # Count the number of Calls for later use.
    num_calls = len(Calls)

########## Show the last line I pasted.

    # Open alm_transfer_order_line.csv and extract the individual records into Parts.
    with open("alm_transfer_order_line.csv", "r") as file:
        reader = csv.DictReader(file)
        Parts = list(reader)

    return Calls, Parts, field_name_mapping, num_calls

def clean_phone_number(phone_number: str) -> str:
    """Removes non-numeric characters from phone number and formats as nnn-nnn-nnn x nnnnnnnn"""
    cleaned_number = re.sub(r"[^\d]", "", phone_number)
    if len(cleaned_number) > 10:
        return f"{cleaned_number[:3]}-{cleaned_number[3:6]}-{cleaned_number[6:10]} x {cleaned_number[10:]}"
    else:
        return f"{cleaned_number[:3]}-{cleaned_number[3:6]}-{cleaned_number[6:]}"

def remove_newlines(text: str) -> str:
    """Removes all newline characters from text"""
    return text.replace("\n", " ")

def data_cleanup(Calls):
    for call in Calls:
        for field in ["u_contact_phone", "u_caller_phone2", "u_alt_contact_phone", "u_alternate_contact_phone2"]:
            if call[field]:
                call[field] = clean_phone_number(call[field])
        # call["csz"] = call["u_city"] + ", " + call["u_state"][:2].upper() + " " + call["u_zip_code"][:5]
        for field in call:
            if str(call[field]) == call[field]:
                call[field] = remove_newlines(call[field])
    return Calls

def handle_parts(Calls, Parts):
    for call in Calls:
    # Relate Calls to Parts in a one-to-many relationship where the Calls "number" 
    # field is matched to the Parts "u_service_task_order" field.
        call["parts"] = []
        for part in Parts:
            if part["u_service_task_order"] == call["number"]:
                call["parts"].append(part)

########## Show the last line of code I pasted.

        # For each Part, append a field to Calls using its tracking_number with an
        # added sequence number that resets with each different tracking_number as the fieldname, 
        # and u_part_description as the value. 
        # The result should be a Call record with zero or more part fields with the fieldname 
        # ({tracking_number1}-1, {tracking_number1}-2, {tracking_number2}-1, {tracking_number2}-2, 
        # etc.) each with the value of u_part_description.
        tracking_number_count = {}
        for part in call["parts"]:
            if part["tracking_number"] not in tracking_number_count:
                tracking_number_count[part["tracking_number"]] = 1
            else:
                tracking_number_count[part["tracking_number"]] += 1
            call[f"Part{tracking_number_count[part['tracking_number']]}: {part['tracking_number']}"] = part["u_part_description"]
            call["Depot: " + part["u_shipto_location"]] = ""
    
    return Calls

def open_maps(calls_list, Calls):
    selected_calls = calls_list.curselection()
    addresses = []
    depot_to_address = {"SPRINGFIELD": "700 International Way, Springfield OR 97401",
                        "MEDFORD": "3600 Terminal Spur Road, Medford OR 97502"}
    for i in selected_calls:
        call = Calls[i]
        for part in call["parts"]:
            for depot in depot_to_address:
                if depot in part["u_shipto_location"]:
                    if depot_to_address[depot] not in addresses:
                        addresses.append(depot_to_address[depot])
                    break
        address = call["u_street_address1"]
        if call["u_street_address2"]:
            address += "+" + call["u_street_address2"]
        address += "+" + call["u_city"] + ",+" + call["u_state"].upper()[:2] + ",+" + call["u_zip_code"][:5]
        address = address.replace(" ","+")
        address = address.replace("++","+")
        addresses.append(address)
    # create the url with addresses
    url = "https://www.google.com/maps/dir/5541+NE+72nd+Ave+Portland+OR+97218/" + "/".join(addresses)
    webbrowser.open(url)

########## Show the last line I pasted.

def get_depot_from_shipto_location(shipto_location, depot_to_address):
    return next(
        (
            address
            for depot, address in depot_to_address.items()
            if depot in shipto_location
        ),
        "Depot not found",
    )

def open_fedex_tracking(Calls):
    tracking_numbers = set()
    for i in range(len(Calls)):
        call = Calls[i]
        for key in call:
            if key.startswith("Part"):
                tracking_numbers.add(key.split(":")[1].strip())

    # create the url with tracking_numbers
    if tracking_numbers:
        url = "https://www.fedex.com/apps/fedextrack/?tracknumbers=" + ",".join(tracking_numbers)
        webbrowser.open(url)

sort_by_number = False

def update_calls_label(calls_label, selected_calls, Calls):
    if len(selected_calls) == 1:
        calls_label.config(text=f"Call {selected_calls[0]+1} of {len(Calls)} calls")
    else:
        calls_label.config(text=f"{len(selected_calls)} calls selected of {len(Calls)} calls")

########## Show only the last line of code I pasted.

def update_preview(gui, preview, selected_calls, Calls, canned_lines, field_name_mapping):
    preview.config(state=NORMAL)
    preview.delete(1.0, END)

    if selected_calls:
        for i in selected_calls:
            call = Calls[i]
            for field in call:
                if call[field] and field != "parts":
                    if field.startswith("Part"):
                        preview.insert(END, field + ":\t")
                    else:
                        preview.insert(END, field_name_mapping[field] + ":\t")
                    preview.insert(END, str(call[field]) + "\n")
            if len(selected_calls)>1:
                preview.insert(END, '_'*20+'\n\n')
    preview.config(state=DISABLED)

def update_calls_list(calls_list, Calls, bool_assigned, bool_pending, bool_complete, sort_by_number):
    calls_list.delete(0, END)
    filtered_calls = [call for call in Calls if (call['state'] == 'Assigned' and bool_assigned.get()) or (call['state'] == 'Pending Dispatch' and bool_pending.get()) or (call['state'] == 'Work Complete' and bool_complete.get())]

    if sort_by_number:
        filtered_calls.sort(key=lambda x: x['number'])

    for i, call in enumerate(filtered_calls):
        calls_list.insert(END, call['number'])

        if call['state'] == 'Assigned':
            calls_list.itemconfig(i, bg='lightgreen')
        elif call['state'] == 'Pending Dispatch':
            calls_list.itemconfig(i, bg='yellow')
        elif call['state'] == 'Work Complete':
            calls_list.itemconfig(i, bg='lightblue')

def update_calls(gui, calls_list, calls_label, preview, Calls, field_name_mapping, bool_assigned, bool_pending, bool_complete, sort_by_number, ctrl_clicked=False):
 
########## Show only the last line of code I pasted.

    selected_calls = list(calls_list.curselection())
    if not ctrl_clicked:
        update_calls_label(calls_label, selected_calls, Calls)
        update_preview(gui, preview, selected_calls, Calls, gui.canned_lines, field_name_mapping)
        update_calls_list(calls_list, Calls, bool_assigned, bool_pending, bool_complete, sort_by_number)
        if selected_calls:
            calls_list.activate(selected_calls[0])
            calls_list.selection_clear(0, END)
            for i in selected_calls:
                calls_list.activate(i)
                calls_list.selection_set(i, last=i)
        if len(selected_calls)>1:
            gui.email_button.config(state="disabled")
        else:
            gui.email_button.config(state="normal")

def toggle_sort_calls_list(gui, calls_list, calls_label, preview, Calls, field_name_mapping, sort_button, bool_assigned, bool_pending, bool_complete):
    global sort_by_number
    sort_by_number = not sort_by_number
    if sort_button["text"] == '''Assign''':
       sort_button["text"] = '''Number'''
    else: 
       sort_button["text"] = '''Assign'''

    # save the selected call numbers before updating
    selected_calls = [Calls[i]['number'] for i in calls_list.curselection()]
        
    update_calls(gui, calls_list, calls_label, preview, Calls, field_name_mapping, bool_assigned, bool_pending, bool_complete, sort_by_number)

    # re-highlight the selected call numbers after updating
    for i, call in enumerate(Calls):
        if call['number'] in selected_calls:
#            calls_list.activate(i)
#            calls_list.selection_set(i, last=i)

            calls_list.activate(i)
            calls_list.select_set(i)
        calls_list.event_generate("<<ListboxSelect>>")

########## Show only the last line of code I pasted.

#    update_calls(calls_list, calls_label, preview, Calls, field_name_mapping, bool_assigned, bool_pending, bool_complete, sort_by_number)
global store
store = {
    'ListBoxFrame': {
        'relx': 0.023,
        'rely': 0.05,
        'relheight': 0.87,
        'relwidth': 0.30,
    },
    'CannedLinesFrame': {
        'relx': 0.023,
        'rely': 0.05,
        'relheight': 0.90,
        'relwidth': 0.30,
    },
    'DropDownsFrame': {
        'relx': 0.353,
        'rely': 0.70,
        'relheight': 0.15,
        'relwidth': 0.25,
    },
}

def toggle_frame(gui):
    global store

    if 'ListBoxFrame' not in store:
        store['ListBoxFrame'] = gui.ListBoxFrame.place_info()
        if gui.CannedLinesFrame.winfo_ismapped():
            store['CannedLinesFrame'] = gui.CannedLinesFrame.place_info()
        if gui.DropDownsFrame.winfo_ismapped():
            store['DropDownsFrame'] = gui.DropDownsFrame.place_info()

    if gui.ListBoxFrame.winfo_ismapped():
        gui.ListBoxFrame.place_forget()
        if 'CannedLinesFrame' in store:
            gui.CannedLinesFrame.place(store['CannedLinesFrame'])
        if 'DropDownsFrame' in store:
            gui.DropDownsFrame.place(store['DropDownsFrame'])
        gui.email_button["text"] = '''Calls'''
        gui.daysheet_button["text"] = '''Generate\nEmail'''
        gui.preview.config(state=NORMAL)
        gui.preview.delete(1.0, END)
    else:
        if 'CannedLinesFrame' in store:
            gui.CannedLinesFrame.place_forget()
        gui.ListBoxFrame.place(store['ListBoxFrame'])
        if 'DropDownsFrame' in store:
            gui.DropDownsFrame.place_forget()
        gui.email_button["text"] = '''Email'''
        gui.daysheet_button["text"] = '''Generate\nDaysheet'''

########## Show only the last of code line I pasted.

def setup(gui, root, DropDownsFrame):
    # Set up the repair type dropdown menu
    repair_types = ["Computer", "Server"]
    gui.repair_type["values"] = repair_types

    # Set up the delay type dropdown menu
    delay_types = ["None", "Delay", "Part Delay", "Long Delay"]
    gui.delay_type.set("None")
    gui.delay_type["values"] = delay_types

    # Set up the start time and end time dropdown menus
    times = [str(x) for x in range(1, 13)]
    gui.starttime["values"] = times
    gui.endtime["values"] = times

    # Set up the repair type dropdown menu
    parts_types = ["System Board", "Keyboard", "LCD Panel", "DC-in cable"]
    gui.parts["values"] = parts_types

    gui.repair_type.current(0)
    gui.delay_type.current(0)
    gui.starttime.current(9)
    gui.endtime.current(11)
    gui.parts.current(0)

def run_gui(gui, Calls, calls_list, preview, calls_label, field_name_mapping):

        calls_list.config(selectmode = "extended")
        gui.maps_button.config(command=lambda: open_maps(gui.calls_list, Calls))
        gui.fedex_button.config(command=lambda: open_fedex_tracking(Calls))
        gui.email_button.config(command=lambda: toggle_frame(gui))

        gui.assn_chkbtn.configure(command=lambda: update_calls(gui, gui.calls_list, gui.calls_label, gui.preview, Calls, field_name_mapping, gui.bool_assigned, gui.bool_pending, gui.bool_complete, sort_by_number))
        gui.pndg_chkbtn.configure(command=lambda: update_calls(gui, gui.calls_list, gui.calls_label, gui.preview, Calls, field_name_mapping, gui.bool_assigned, gui.bool_pending, gui.bool_complete, sort_by_number))
        gui.comp_chkbtn.configure(command=lambda: update_calls(gui, gui.calls_list, gui.calls_label, gui.preview, Calls, field_name_mapping, gui.bool_assigned, gui.bool_pending, gui.bool_complete, sort_by_number))
        gui.sort_chkbtn.config(command=lambda: toggle_sort_calls_list(gui, calls_list, calls_label, preview, Calls, field_name_mapping, gui.sort_chkbtn, gui.bool_assigned, gui.bool_pending, gui.bool_complete))

        # Create a right-click menu for the values_textbox widget that includes a "Copy" option and "Select All" option
        menu = Menu(preview, tearoff=0)
        menu.add_command(label="Copy", accelerator="Ctrl+C", command=lambda: pyperclip.copy(gui.preview.get(SEL_FIRST, SEL_LAST)))
        menu.add_command(label="Select All", accelerator="Ctrl+A", command=lambda: gui.preview.tag_add(SEL, "1.0", END))
        preview.bind("<Button-3>", lambda event: menu.post(event.x_root, event.y_root))

########## Show only the last lineof code I pasted.

        return gui

def on_select(gui, calls_list, preview, Calls, field_name_mapping):
    selected = None
    with contextlib.suppress(Exception):
        selected = calls_list.get(calls_list.curselection())
        preview.delete(1.0, END)
        preview.insert(END, selected)
        gui.email_button.config(state="disabled")

    for call in Calls:

        if call["number"] == selected:
            preview.config(state=NORMAL)
            preview.delete(1.0, END)
            for field_name in field_name_mapping.values():
                text_widget = getattr(gui, field_name.lower())
                text_widget.delete(0, END)

            for field in call:
                if call[field] and field != "parts":
                    if field.startswith("Part"):
                        preview.insert(END, field + ":\t")
                    else:
                        preview.insert(END, field_name_mapping[field] + ":\t")
                        text_widget = getattr(gui, field_name_mapping[field].lower())
                        text_widget.delete(0, END)
                        text_widget.insert(0, str(call[field]))
                    preview.insert(END, str(call[field]) + "\n")
            preview.config(state=NORMAL)


# Create a flag to check if the list is sorted or not
is_sorted = False

########## Show only the last line of code I pasted.

def main(*args):  # sourcery skip: boolean-if-exp-identity
    '''Main entry point for the application.'''
    global root
    root = tk.Tk()
    root.protocol( 'WM_DELETE_WINDOW' , root.destroy)
    source = os.path.join(os.environ.get('HOMEPATH'), 'Downloads')
    Calls, Parts, field_name_mapping, num_calls = file_handling(source)

    Calls = handle_parts(Calls, Parts)
    
    gui = ParserGUI.ParserGUI(root)
    setup(gui, root, gui.DropDownsFrame)
    gui = run_gui(gui, Calls, gui.calls_list, gui.preview, gui.calls_label, field_name_mapping)
    gui.assn_chkbtn.select()
    gui.pndg_chkbtn.select()
    gui.comp_chkbtn.select()

    update_calls(gui, gui.calls_list, gui.calls_label, gui.preview, Calls, field_name_mapping, gui.bool_assigned, gui.bool_pending, gui.bool_complete, sort_by_number)
    gui.calls_list.bind('<<ListboxSelect>>', lambda _: on_select(gui, gui.calls_list, gui.preview, Calls, field_name_mapping))
    gui.calls_list.bind("<Button-1>", lambda _: update_calls(gui, gui.calls_list, gui.calls_label, gui.preview, Calls, field_name_mapping, gui.bool_assigned, gui.bool_pending, gui.bool_complete, sort_by_number))
    gui.calls_list.bind("<Control-Button-1>", lambda event: update_calls(gui, gui.calls_list, gui.calls_label, gui.preview, Calls, field_name_mapping, gui.bool_assigned, gui.bool_pending, gui.bool_complete,sort_by_number, ctrl_clicked=True if event.state & 0x0004 else False))
    gui.calls_list.bind("<ButtonRelease-1>", lambda _: update_calls(gui, gui.calls_list, gui.calls_label, gui.preview, Calls, field_name_mapping, gui.bool_assigned, gui.bool_pending, gui.bool_complete, sort_by_number))

    mainloop()

if __name__ == '__main__':
    main()

########## This is SvcNowParser.py. Do you have all the code?
