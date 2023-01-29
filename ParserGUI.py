import sys
import tkinter as tk
import tkinter.ttk as ttk
from tkinter.constants import *
import os.path
import SvcNowParser_SUCCESS1

_script = sys.argv[0]
_location = os.path.dirname(_script)

_bgcolor = '#d9d9d9'  # X11 color: 'gray85'
_fgcolor = '#000000'  # X11 color: 'black'
_compcolor = 'gray40' # X11 color: #666666
_ana1color = '#c3c3c3' # Closest X11 color: 'gray76'
_ana2color = 'beige' # X11 color: #f5f5dc
_tabfg1 = 'black' 
_tabfg2 = 'black' 
_tabbg1 = 'grey75' 
_tabbg2 = 'grey89' 
_bgmode = 'light' 

class ParserGUI:
    def __init__(self, top=None):
        '''This class configures and populates the toplevel window.
           top is the toplevel containing window.'''

        top.geometry("1312x791+106+12")
        top.minsize(120, 1)
        top.maxsize(3460, 1061)
        top.resizable(1,  1)
        top.title("ServiceNow Parser")
        top.configure(background="#d9d9d9")
        top.configure(highlightbackground="#d9d9d9")
        top.configure(highlightcolor="black")

        self.top = top
        self.bool_sorts = tk.BooleanVar()
        self.bool_assigned = tk.BooleanVar()
        self.bool_pending = tk.BooleanVar()
        self.bool_complete = tk.BooleanVar()
        self.canned_lines = []

########## Show the last line of code I pasted.
##### CallsList Frame

        self.ListBoxFrame = tk.Frame(self.top)

        self.calls_label = tk.Label(self.ListBoxFrame)
        self.calls_label.configure(activebackground="#f9f9f9")
        self.calls_label.configure(anchor='w')
        self.calls_label.configure(background="#d9d9d9")
        self.calls_label.configure(compound='left')
        self.calls_label.configure(disabledforeground="#a3a3a3")
        self.calls_label.configure(foreground="#000000")
        self.calls_label.configure(highlightbackground="#d9d9d9")
        self.calls_label.configure(highlightcolor="black")
        self.calls_label.configure(text='''Calls''')

        self.calls_list = tk.Listbox(self.ListBoxFrame)
        self.calls_list.configure(background="white")
        self.calls_list.configure(disabledforeground="#a3a3a3")
        self.calls_list.configure(font="-family {Courier New} -size 10")
        self.calls_list.configure(foreground="#000000")
        self.calls_list.configure(highlightbackground="#d9d9d9")
        self.calls_list.configure(highlightcolor="black")
        self.calls_list.configure(selectbackground="#c4c4c4")
        self.calls_list.configure(selectforeground="black")

        self.sort_chkbtn = tk.Checkbutton(self.ListBoxFrame)
        self.sort_chkbtn.configure(activebackground="beige")
        self.sort_chkbtn.configure(activeforeground="black")
        self.sort_chkbtn.configure(background="#d9d9d9")
        self.sort_chkbtn.configure(compound='left')
        self.sort_chkbtn.configure(disabledforeground="#a3a3a3")
        self.sort_chkbtn.configure(foreground="#000000")
        self.sort_chkbtn.configure(highlightbackground="#d9d9d9")
        self.sort_chkbtn.configure(highlightcolor="black")
        self.sort_chkbtn.configure(pady="0")
        self.sort_chkbtn.configure(text='''Sort by Number''')
        self.sort_chkbtn.configure(variable=self.bool_sorts)

########## Show the last line of code I pasted.

        self.assn_chkbtn = tk.Checkbutton(self.ListBoxFrame)
        self.assn_chkbtn.select()
        self.assn_chkbtn.configure(activebackground="beige")
        self.assn_chkbtn.configure(activeforeground="black")
        self.assn_chkbtn.configure(anchor='w')
        self.assn_chkbtn.configure(background="#d9d9d9")
        self.assn_chkbtn.configure(compound='left')
        self.assn_chkbtn.configure(disabledforeground="#a3a3a3")
        self.assn_chkbtn.configure(foreground="#000000")
        self.assn_chkbtn.configure(highlightbackground="#d9d9d9")
        self.assn_chkbtn.configure(highlightcolor="black")
        self.assn_chkbtn.configure(justify='left')
        self.assn_chkbtn.configure(selectcolor="#d9d9d9")
        self.assn_chkbtn.configure(text='''Assigned''')
        self.assn_chkbtn.configure(variable=self.bool_assigned)
        
        self.pndg_chkbtn = tk.Checkbutton(self.ListBoxFrame)
        self.pndg_chkbtn.select()
        self.pndg_chkbtn.configure(activebackground="beige")
        self.pndg_chkbtn.configure(activeforeground="black")
        self.pndg_chkbtn.configure(anchor='w')
        self.pndg_chkbtn.configure(background="#d9d9d9")
        self.pndg_chkbtn.configure(compound='left')
        self.pndg_chkbtn.configure(disabledforeground="#a3a3a3")
        self.pndg_chkbtn.configure(foreground="#000000")
        self.pndg_chkbtn.configure(highlightbackground="#d9d9d9")
        self.pndg_chkbtn.configure(highlightcolor="black")
        self.pndg_chkbtn.configure(justify='left')
        self.pndg_chkbtn.configure(selectcolor="#d9d9d9")
        self.pndg_chkbtn.configure(text='''Pending''')
        self.pndg_chkbtn.configure(variable=self.bool_pending)
        
########## Show the last line of code I pasted.

        self.comp_chkbtn = tk.Checkbutton(self.ListBoxFrame)
        self.comp_chkbtn.select()
        self.comp_chkbtn.configure(activebackground="beige")
        self.comp_chkbtn.configure(activeforeground="black")
        self.comp_chkbtn.configure(anchor='w')
        self.comp_chkbtn.configure(background="#d9d9d9")
        self.comp_chkbtn.configure(compound='left')
        self.comp_chkbtn.configure(disabledforeground="#a3a3a3")
        self.comp_chkbtn.configure(foreground="#000000")
        self.comp_chkbtn.configure(highlightbackground="#d9d9d9")
        self.comp_chkbtn.configure(highlightcolor="black")
        self.comp_chkbtn.configure(justify='left')
        self.comp_chkbtn.configure(selectcolor="#d9d9d9")
        self.comp_chkbtn.configure(text='''Complete''')
        self.comp_chkbtn.configure(variable=self.bool_complete)

### Place for CallsList frame

        self.calls_label.place(relx=0.00, rely=0.00, height=21, relwidth=0.999)
        self.calls_list.place(relx=0.0, rely=0.028, relheight=.95, relwidth=0.999)
        self.sort_chkbtn.place(relx=0.0, rely=0.95, height=44, width=120)
        self.assn_chkbtn.place(relx=0.28, rely=0.95, height=44, width=120)
        self.pndg_chkbtn.place(relx=0.50, rely=0.95, height=44, width=120)
        self.comp_chkbtn.place(relx=0.75, rely=0.95, height=44, width=120)
        self.ListBoxFrame.place(relx=0.023, rely=0.05, relheight=0.87, relwidth=0.30)
        
### Preview Window

        self.preview = tk.Text(self.top)
        self.preview.place(relx=0.625, rely=0.076, relheight=0.789, relwidth=0.343)
        self.preview.configure(background="white")
        self.preview.configure(font="-family {Courier New} -size 10")
        self.preview.configure(foreground="black")
        self.preview.configure(highlightbackground="#d9d9d9")
        self.preview.configure(highlightcolor="black")
        self.preview.configure(insertbackground="black")
        self.preview.configure(selectbackground="#c4c4c4")
        self.preview.configure(selectforeground="black")
        self.preview.configure(wrap="word")
        self.preview.insert(1.0, "")  # insert an empty string to make the widget blank

########## Show the last line of code I pasted.

        self.fields_values_label = tk.Label(self.top)
        self.fields_values_label.place(relx=0.625, rely=0.050, height=20, width=114)
        self.fields_values_label.configure(activebackground="#f9f9f9")
        self.fields_values_label.configure(anchor='w')
        self.fields_values_label.configure(background="#d9d9d9")
        self.fields_values_label.configure(compound='left')
        self.fields_values_label.configure(disabledforeground="#a3a3a3")
        self.fields_values_label.configure(foreground="#000000")
        self.fields_values_label.configure(highlightbackground="#d9d9d9")
        self.fields_values_label.configure(highlightcolor="black")
        self.fields_values_label.configure(text='''DaySheet Preview''')

##### CannedLines Frame

        self.CannedLinesFrame = tk.Frame(self.top)

        self.bg_label = tk.Label(self.CannedLinesFrame)
        self.bg_label.configure(anchor='w')
        self.bg_label.configure(background="#d9d9d9")
        self.bg_label.configure(compound='left')
        self.bg_label.configure(disabledforeground="#a3a3a3")
        self.bg_label.configure(foreground="#000000")

        self.preset1_button = tk.Button(self.CannedLinesFrame)
        self.preset1_button.configure(activebackground="beige")
        self.preset1_button.configure(activeforeground="black")
        self.preset1_button.configure(background="#d9d9d9")
        self.preset1_button.configure(compound='left')
        self.preset1_button.configure(disabledforeground="#a3a3a3")
        self.preset1_button.configure(foreground="#000000")
        self.preset1_button.configure(highlightbackground="#d9d9d9")
        self.preset1_button.configure(highlightcolor="black")
        self.preset1_button.configure(pady="0")
        self.preset1_button.configure(text='''1st Email\nPreset''')

        self.preset2_button = tk.Button(self.CannedLinesFrame)
        self.preset2_button.configure(activebackground="beige")
        self.preset2_button.configure(activeforeground="black")
        self.preset2_button.configure(background="#d9d9d9")
        self.preset2_button.configure(compound='left')
        self.preset2_button.configure(disabledforeground="#a3a3a3")
        self.preset2_button.configure(foreground="#000000")
        self.preset2_button.configure(highlightbackground="#d9d9d9")
        self.preset2_button.configure(highlightcolor="black")
        self.preset2_button.configure(pady="0")
        self.preset2_button.configure(text='''2nd Email\nPreset''')
 
########## Show the last line of code I pasted.

        self.preset3_button = tk.Button(self.CannedLinesFrame)
        self.preset3_button.configure(activebackground="beige")
        self.preset3_button.configure(activeforeground="black")
        self.preset3_button.configure(background="#d9d9d9")
        self.preset3_button.configure(compound='left')
        self.preset3_button.configure(disabledforeground="#a3a3a3")
        self.preset3_button.configure(foreground="#000000")
        self.preset3_button.configure(highlightbackground="#d9d9d9")
        self.preset3_button.configure(highlightcolor="black")
        self.preset3_button.configure(pady="0")
        self.preset3_button.configure(text='''3rd Email\nPreset''')
 
        self.preset4_button = tk.Button(self.CannedLinesFrame)
        self.preset4_button.configure(activebackground="beige")
        self.preset4_button.configure(activeforeground="black")
        self.preset4_button.configure(background="#d9d9d9")
        self.preset4_button.configure(compound='left')
        self.preset4_button.configure(disabledforeground="#a3a3a3")
        self.preset4_button.configure(foreground="#000000")
        self.preset4_button.configure(highlightbackground="#d9d9d9")
        self.preset4_button.configure(highlightcolor="black")
        self.preset4_button.configure(pady="0")
        self.preset4_button.configure(text='''4th Email\nPreset''')

        self.canned_lines_label = tk.Label(self.CannedLinesFrame)
        self.canned_lines_label.configure(anchor='w')
        self.canned_lines_label.configure(background="#d9d9d9")
        self.canned_lines_label.configure(compound='left')
        self.canned_lines_label.configure(disabledforeground="#a3a3a3")
        self.canned_lines_label.configure(foreground="#000000")
        self.canned_lines_label.configure(text='''Select Lines:''')

        canned_lines_field = tk.Text(self.CannedLinesFrame)
        canned_lines_field.configure(background="white")
        canned_lines_field.configure(cursor="xterm")
        canned_lines_field.configure(font="-family {Courier New} -size 10")
        canned_lines_field.configure(foreground="black")
        canned_lines_field.configure(highlightbackground="#d9d9d9")
        canned_lines_field.configure(highlightcolor="#d9d9d9")
        canned_lines_field.configure(selectbackground="#c4c4c4")
        canned_lines_field.configure(selectforeground="black")

########## Show the last line of code I pasted.

        # canned_lines_field.trace("w", lambda name, index, mode, var=self.name: update_preview(self))
        def add_line_to_preview(self, var, line):
            if var.get():
                self.preview.config(state=tk.NORMAL)
                self.preview.insert(tk.END, line + "\n")
                self.preview.config(state=tk.DISABLED)

        def remove_line_from_preview(self, var, line):
            if not var.get():
                self.preview.config(state=tk.NORMAL)
                self.preview.delete(line + "\n")
                self.preview.config(state=tk.DISABLED)

        scrollbar = tk.Scrollbar(canned_lines_field)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        var = tk.IntVar()
        vars = []
        with open('canned_lines.txt', 'r') as f:
            self.canned_lines = f.readlines()
            for line in self.canned_lines:
                vars.append(line.strip())
                var = tk.IntVar()
                vars.append(var)
                checkbutton = tk.Checkbutton(canned_lines_field, variable=var, command=lambda var=var, line=line: add_line_to_preview(self, var, line))
                var.trace("w", lambda *args, var=var, line=line: remove_line_from_preview(self, var, line))

                canned_lines_field.window_create("end", window=checkbutton)
                canned_lines_field.insert("end", line+"\n")
        canned_lines_field.config(yscrollcommand=scrollbar.set)
        scrollbar.config(command=canned_lines_field.yview)

    # disable the widget so users can't insert text into it
        canned_lines_field.configure(state="disabled")     

        self.canned_lines_label.place(relx=0.00, rely=0.00, height=20, relwidth=0.999)
        self.bg_label.place(relx=0, rely=0.915, height=62, relwidth=0.999)
        self.preset1_button.place(relx=0.00, rely=0.93, height=44, width=75)
        self.preset2_button.place(relx=0.27, rely=0.93, height=44, width=75)
        self.preset3_button.place(relx=0.54, rely=0.93, height=44, width=75)
        self.preset4_button.place(relx=0.81, rely=0.93, height=44, width=75)
        canned_lines_field.place(relx=0, rely=0.024, relheight=0.90, relwidth=0.999)
 
########## Show the last line of code I pasted.
##### Buttons

        self.daysheet_button = tk.Button(self.top)
        self.daysheet_button.configure(activebackground="beige")
        self.daysheet_button.configure(activeforeground="black")
        self.daysheet_button.configure(background="#d9d9d9")
        self.daysheet_button.configure(compound='left')
        self.daysheet_button.configure(disabledforeground="#a3a3a3")
        self.daysheet_button.configure(foreground="#000000")
        self.daysheet_button.configure(highlightbackground="#d9d9d9")
        self.daysheet_button.configure(highlightcolor="black")
        self.daysheet_button.configure(pady="0")
        self.daysheet_button.configure(text='''Generate\nDaysheet''')

        self.maps_button = tk.Button(self.top)
        self.maps_button.configure(activebackground="beige")
        self.maps_button.configure(activeforeground="black")
        self.maps_button.configure(background="#d9d9d9")
        self.maps_button.configure(compound='left')
        self.maps_button.configure(disabledforeground="#a3a3a3")
        self.maps_button.configure(foreground="#000000")
        self.maps_button.configure(highlightbackground="#d9d9d9")
        self.maps_button.configure(highlightcolor="black")
        self.maps_button.configure(pady="0")
        self.maps_button.configure(text='''Google\nMap''')

        self.fedex_button = tk.Button(self.top)
        self.fedex_button.configure(activebackground="beige")
        self.fedex_button.configure(activeforeground="black")
        self.fedex_button.configure(background="#d9d9d9")
        self.fedex_button.configure(compound='left')
        self.fedex_button.configure(disabledforeground="#a3a3a3")
        self.fedex_button.configure(foreground="#000000")
        self.fedex_button.configure(highlightbackground="#d9d9d9")
        self.fedex_button.configure(highlightcolor="black")
        self.fedex_button.configure(pady="0")
        self.fedex_button.configure(text='''FedEx\nTracking''')

        self.email_button = tk.Button(self.top)
        self.email_button.configure(activebackground="beige")
        self.email_button.configure(activeforeground="black")
        self.email_button.configure(background="#d9d9d9")
        self.email_button.configure(compound='left')
        self.email_button.configure(disabledforeground="#a3a3a3")
        self.email_button.configure(foreground="#000000")
        self.email_button.configure(highlightbackground="#d9d9d9")
        self.email_button.configure(highlightcolor="black")
        self.email_button.configure(pady="0")
        self.email_button.configure(text='''Email''')

########## Show the last line of code I pasted.
###### Place for buttons

        self.email_button.place(relx=0.35, rely=0.885, height=45, width=90)
        self.maps_button.place(relx=0.45, rely=0.885, height=45, width=90)
        self.fedex_button.place(relx=0.55, rely=0.885, height=45, width=90)
        self.daysheet_button.place(relx=0.75, rely=0.885, height=45, width=90)

##### Parsed Fields

        self.number = tk.Entry(self.top)
        self.number.configure(background="white")
        self.number.configure(disabledforeground="#a3a3a3")
        self.number.configure(font="-family {Segoe UI} -size 10")
        self.number.configure(foreground="#000000")
        self.number.configure(insertbackground="black")        

        self.dspnum = tk.Entry(self.top)
        self.dspnum.configure(background="white")
        self.dspnum.configure(disabledforeground="#a3a3a3")
        self.dspnum.configure(font="-family {Segoe UI} -size 10")
        self.dspnum.configure(foreground="#000000")
        self.dspnum.configure(insertbackground="black")        

        self.name = tk.Entry(self.top)
        self.name.configure(background="white")
        self.name.configure(disabledforeground="#a3a3a3")
        self.name.configure(font="-family {Segoe UI} -size 10")
        self.name.configure(foreground="#000000")
        self.name.configure(insertbackground="black")

        self.comp = tk.Entry(self.top)
        self.comp.configure(background="white")
        self.comp.configure(disabledforeground="#a3a3a3")
        self.comp.configure(font="-family {Segoe UI} -size 10")
        self.comp.configure(foreground="#000000")
        self.comp.configure(insertbackground="black")        

        self.status = tk.Entry(self.top)
        self.status.configure(background="white")
        self.status.configure(disabledforeground="#a3a3a3")
        self.status.configure(font="-family {Segoe UI} -size 10")
        self.status.configure(foreground="#000000")
        self.status.configure(insertbackground="black")
 
########## Show the last line of code I pasted.

        self.addr1 = tk.Entry(self.top)
        self.addr1.configure(background="white")
        self.addr1.configure(disabledforeground="#a3a3a3")
        self.addr1.configure(font="-family {Segoe UI} -size 10")
        self.addr1.configure(foreground="#000000")
        self.addr1.configure(insertbackground="black")

        self.addr2 = tk.Entry(self.top)
        self.addr2.configure(background="white")
        self.addr2.configure(disabledforeground="#a3a3a3")
        self.addr2.configure(font="-family {Segoe UI} -size 10")
        self.addr2.configure(foreground="#000000")
        self.addr2.configure(insertbackground="black")

        self.city = tk.Entry(self.top)
        self.city.configure(background="white")
        self.city.configure(disabledforeground="#a3a3a3")
        self.city.configure(font="-family {Segoe UI} -size 10")
        self.city.configure(foreground="#000000")
        self.city.configure(insertbackground="black")

        self.state = tk.Entry(self.top)
        self.state.configure(background="white")
        self.state.configure(disabledforeground="#a3a3a3")
        self.state.configure(font="-family {Segoe UI} -size 10")
        self.state.configure(foreground="#000000")
        self.state.configure(insertbackground="black")

        self.zip = tk.Entry(self.top)
        self.zip.configure(background="white")
        self.zip.configure(disabledforeground="#a3a3a3")
        self.zip.configure(font="-family {Segoe UI} -size 10")
        self.zip.configure(foreground="#000000")
        self.zip.configure(insertbackground="black")

        self.phone = tk.Entry(self.top)
        self.phone.configure(background="white")
        self.phone.configure(disabledforeground="#a3a3a3")
        self.phone.configure(font="-family {Segoe UI} -size 10")
        self.phone.configure(foreground="#000000")
        self.phone.configure(insertbackground="black")

########## Show the last line of code I pasted.

        self.phone2 = tk.Entry(self.top)
        self.phone2.configure(background="white")
        self.phone2.configure(disabledforeground="#a3a3a3")
        self.phone2.configure(font="-family {Segoe UI} -size 10")
        self.phone2.configure(foreground="#000000")
        self.phone2.configure(insertbackground="black")

        self.email = tk.Entry(self.top)
        self.email.configure(background="white")
        self.email.configure(disabledforeground="#a3a3a3")
        self.email.configure(font="-family {Segoe UI} -size 10")
        self.email.configure(foreground="#000000")
        self.email.configure(insertbackground="black")

        self.altcontact = tk.Entry(self.top)
        self.altcontact.configure(background="white")
        self.altcontact.configure(disabledforeground="#a3a3a3")
        self.altcontact.configure(font="-family {Segoe UI} -size 10")
        self.altcontact.configure(foreground="#000000")
        self.altcontact.configure(insertbackground="black")

        self.altphone = tk.Entry(self.top)
        self.altphone.configure(background="white")
        self.altphone.configure(disabledforeground="#a3a3a3")
        self.altphone.configure(font="-family {Segoe UI} -size 10")
        self.altphone.configure(foreground="#000000")
        self.altphone.configure(insertbackground="black")

        self.altphone2 = tk.Entry(self.top)
        self.altphone2.configure(background="white")
        self.altphone2.configure(disabledforeground="#a3a3a3")
        self.altphone2.configure(font="-family {Segoe UI} -size 10")
        self.altphone2.configure(foreground="#000000")
        self.altphone2.configure(insertbackground="black")

        self.altemail = tk.Entry(self.top)
        self.altemail.configure(background="white")
        self.altemail.configure(disabledforeground="#a3a3a3")
        self.altemail.configure(font="-family {Segoe UI} -size 10")
        self.altemail.configure(foreground="#000000")
        self.altemail.configure(insertbackground="black")

########## Show the last line of code I pasted.

        self.model = tk.Entry(self.top)
        self.model.configure(background="white")
        self.model.configure(disabledforeground="#a3a3a3")
        self.model.configure(font="-family {Segoe UI} -size 10")
        self.model.configure(foreground="#000000")
        self.model.configure(insertbackground="black")
        
        self.serial = tk.Entry(self.top)
        self.serial.configure(background="white")
        self.serial.configure(disabledforeground="#a3a3a3")
        self.serial.configure(font="-family {Segoe UI} -size 10")
        self.serial.configure(foreground="#000000")
        self.serial.configure(insertbackground="black")
        
        self.issue = tk.Entry(self.top)
        self.issue.configure(background="white")
        self.issue.configure(disabledforeground="#a3a3a3")
        self.issue.configure(font="-family {Segoe UI} -size 10")
        self.issue.configure(foreground="#000000")
        self.issue.configure(insertbackground="black")
                
        self.notes = tk.Entry(self.top)
        self.notes.configure(background="white")
        self.notes.configure(disabledforeground="#a3a3a3")
        self.notes.configure(font="-family {Segoe UI} -size 10")
        self.notes.configure(foreground="#000000")
        self.notes.configure(insertbackground="black")

########## Show the last line of code I pasted.
##### Place for Parsed Field

        #self.status.place(x=680, y=470, height=20, width=84)
        self.number.place(relx=0.365, rely=0.075, height=20, width=125)
        self.dspnum.place(relx=0.500, rely=0.075, height=20, width=125)
        self.name.place(relx=0.365, rely=0.140, height=20, width=300)
        self.comp.place(relx=0.365, rely=0.205, height=20, width=300)
        self.addr1.place(relx=0.365, rely=0.270, height=20, width=300)
        self.addr2.place(relx=0.365, rely=0.335, height=20, width=300)
        self.city.place(relx=0.365, rely=0.400, height=20, width=180)
        self.state.place(relx=0.520, rely=0.400, height=20, width=25)
        self.zip.place(relx=0.555, rely=0.400, height=20, width=50)
        self.phone.place(relx=0.365, rely=0.465, height=20, width=80)
        self.phone2.place(relx=0.435, rely=0.465, height=20, width=80)
        self.email.place(relx=0.505, rely=0.465, height=20, width=115)
        self.altcontact.place(relx=0.365, rely=0.530, height=20, width=84)
        self.altphone.place(relx=0.365, rely=0.595, height=20, width=80)
        self.altphone2.place(relx=0.435, rely=0.595, height=20, width=80)
        self.altemail.place(relx=0.505, rely=0.595, height=20, width=115)
        self.model.place(relx=0.365, rely=0.660, height=20, width=80)
        self.serial.place(relx=0.435, rely=0.660, height=20, width=80)
        # self.issue.place(x=480, y=710, height=20, width=94)
        # self.notes.place(x=680, y=710, height=20, width=94)

##### Dropdowns
        
        self.parts = ttk.Combobox(self.top)
        self.parts_var = tk.StringVar()
        self.parts.configure(textvariable=self.parts_var)
        self.parts.configure(takefocus="")

        self.DropDownsFrame = tk.Frame(self.top)

        self.repair_type = ttk.Combobox(self.DropDownsFrame)
        self.repair_type_var = tk.StringVar()
        self.repair_type.configure(textvariable=self.repair_type_var)
        self.repair_type.configure(takefocus="")
        
        self.delay_type = ttk.Combobox(self.DropDownsFrame)
        self.delay_type_var = tk.StringVar()
        self.delay_type.configure(textvariable=self.delay_type_var)
        self.delay_type.configure(takefocus="")

        self.starttime = ttk.Combobox(self.DropDownsFrame)
        self.starttime_var = tk.StringVar()
        self.starttime.configure(textvariable=self.starttime_var)
        self.starttime.configure(takefocus="")

########## Show the last line of code I pasted.

        self.endtime = ttk.Combobox(self.DropDownsFrame)
        self.endtime_var = tk.StringVar()
        self.endtime.configure(textvariable=self.endtime_var)
        self.endtime.configure(takefocus="")

##### Dropdown labels

        self.parts_field_label = tk.Label(self.top)
        self.parts_field_label.configure(anchor='w')
        self.parts_field_label.configure(background="#d9d9d9")
        self.parts_field_label.configure(compound='left')
        self.parts_field_label.configure(disabledforeground="#a3a3a3")
        self.parts_field_label.configure(foreground="#000000")
        self.parts_field_label.configure(text='''Parts''')

        self.repair_type_label = tk.Label(self.DropDownsFrame)
        self.repair_type_label.configure(anchor='w')
        self.repair_type_label.configure(background="#d9d9d9")
        self.repair_type_label.configure(compound='left')
        self.repair_type_label.configure(disabledforeground="#a3a3a3")
        self.repair_type_label.configure(foreground="#000000")
        self.repair_type_label.configure(text='''Repair''')

        self.delay_type_label = tk.Label(self.DropDownsFrame)
        self.delay_type_label.configure(anchor='w')
        self.delay_type_label.configure(background="#d9d9d9")
        self.delay_type_label.configure(compound='left')
        self.delay_type_label.configure(disabledforeground="#a3a3a3")
        self.delay_type_label.configure(foreground="#000000")
        self.delay_type_label.configure(text='''Delay''')
        
        self.starttime_field_label = tk.Label(self.DropDownsFrame)
        self.starttime_field_label.configure(anchor='w')
        self.starttime_field_label.configure(background="#d9d9d9")
        self.starttime_field_label.configure(compound='left')
        self.starttime_field_label.configure(disabledforeground="#a3a3a3")
        self.starttime_field_label.configure(foreground="#000000")
        self.starttime_field_label.configure(text='''Start Hour''')

        self.endtime_field_label = tk.Label(self.DropDownsFrame)
        self.endtime_field_label.configure(anchor='w')
        self.endtime_field_label.configure(background="#d9d9d9")
        self.endtime_field_label.configure(compound='left')
        self.endtime_field_label.configure(disabledforeground="#a3a3a3")
        self.endtime_field_label.configure(foreground="#000000")
        self.endtime_field_label.configure(text='''End Hour''')

########## Show the last line of code I pasted.
##### Place for dropdowns

        self.parts.place(relx=0.505, rely=0.660, height=20, width=115)

        self.repair_type.place(relx=0.0, rely=0.245, height=20, width=75)
        self.delay_type.place(relx=0.5, rely=0.245, height=20, width=80)
        self.starttime.place(relx=0.0, rely=0.745, height=20, width=43)
        self.endtime.place(relx=0.5, rely=0.745, height=20, width=43)
        self.DropDownsFrame.place(relx=0.353, rely=0.70, relheight=0.15, relwidth=0.25)
        self.DropDownsFrame.place_forget()

##### Place for dropdown labels

        self.parts_field_label.place(relx=0.505, rely=0.635, height=20, width=32)

        self.repair_type_label.place(relx=0.3, rely=0.245, height=20, width=45)
        self.delay_type_label.place(relx=0.85, rely=0.245, height=20, width=34)
        self.starttime_field_label.place(relx=0.2, rely=0.745, height=20, width=60)
        self.endtime_field_label.place(relx=0.7, rely=0.745, height=20, width=56)

###### Labels

        self.callnum_label = tk.Label(self.top)
        self.callnum_label.configure(anchor='w')
        self.callnum_label.configure(background="#d9d9d9")
        self.callnum_label.configure(compound='left')
        self.callnum_label.configure(disabledforeground="#a3a3a3")
        self.callnum_label.configure(foreground="#000000")
        self.callnum_label.configure(text='''Call Number''')

        self.dspnum_label = tk.Label(self.top)
        self.dspnum_label.configure(anchor='w')
        self.dspnum_label.configure(background="#d9d9d9")
        self.dspnum_label.configure(compound='left')
        self.dspnum_label.configure(disabledforeground="#a3a3a3")
        self.dspnum_label.configure(foreground="#000000")
        self.dspnum_label.configure(text='''DSP Number''')

########## Show the last line of code I pasted.

        self.name_field_label = tk.Label(self.top)
        self.name_field_label.configure(anchor='w')
        self.name_field_label.configure(background="#d9d9d9")
        self.name_field_label.configure(compound='left')
        self.name_field_label.configure(disabledforeground="#a3a3a3")
        self.name_field_label.configure(foreground="#000000")
        self.name_field_label.configure(text='''Name''')

        self.comp_field_label = tk.Label(self.top)
        self.comp_field_label.configure(anchor='w')
        self.comp_field_label.configure(background="#d9d9d9")
        self.comp_field_label.configure(compound='left')
        self.comp_field_label.configure(disabledforeground="#a3a3a3")
        self.comp_field_label.configure(foreground="#000000")
        self.comp_field_label.configure(text='''Company''')

        self.addr1_field_label = tk.Label(self.top)
        self.addr1_field_label.configure(anchor='w')
        self.addr1_field_label.configure(background="#d9d9d9")
        self.addr1_field_label.configure(compound='left')
        self.addr1_field_label.configure(disabledforeground="#a3a3a3")
        self.addr1_field_label.configure(foreground="#000000")
        self.addr1_field_label.configure(text='''Address 1''')

        self.addr2_field_label = tk.Label(self.top)
        self.addr2_field_label.configure(anchor='w')
        self.addr2_field_label.configure(background="#d9d9d9")
        self.addr2_field_label.configure(compound='left')
        self.addr2_field_label.configure(disabledforeground="#a3a3a3")
        self.addr2_field_label.configure(foreground="#000000")
        self.addr2_field_label.configure(text='''Address 2''')

        self.city_field_label = tk.Label(self.top)
        self.city_field_label.configure(anchor='w')
        self.city_field_label.configure(background="#d9d9d9")
        self.city_field_label.configure(compound='left')
        self.city_field_label.configure(disabledforeground="#a3a3a3")
        self.city_field_label.configure(foreground="#000000")
        self.city_field_label.configure(text='''City''')

########## Show the last line of code I pasted.

        self.state_field_label = tk.Label(self.top)
        self.state_field_label.configure(anchor='w')
        self.state_field_label.configure(background="#d9d9d9")
        self.state_field_label.configure(compound='left')
        self.state_field_label.configure(disabledforeground="#a3a3a3")
        self.state_field_label.configure(foreground="#000000")
        self.state_field_label.configure(text='''State''')

        self.zip_field_label = tk.Label(self.top)
        self.zip_field_label.configure(anchor='w')
        self.zip_field_label.configure(background="#d9d9d9")
        self.zip_field_label.configure(compound='left')
        self.zip_field_label.configure(disabledforeground="#a3a3a3")
        self.zip_field_label.configure(foreground="#000000")
        self.zip_field_label.configure(text='''Zip''')

        self.phone1_field_label = tk.Label(self.top)
        self.phone1_field_label.configure(anchor='w')
        self.phone1_field_label.configure(background="#d9d9d9")
        self.phone1_field_label.configure(compound='left')
        self.phone1_field_label.configure(disabledforeground="#a3a3a3")
        self.phone1_field_label.configure(foreground="#000000")
        self.phone1_field_label.configure(text='''Phone''')

        self.phone2_field_label = tk.Label(self.top)
        self.phone2_field_label.configure(anchor='w')
        self.phone2_field_label.configure(background="#d9d9d9")
        self.phone2_field_label.configure(compound='left')
        self.phone2_field_label.configure(disabledforeground="#a3a3a3")
        self.phone2_field_label.configure(foreground="#000000")
        self.phone2_field_label.configure(text='''Phone 2''')

        self.email_field_label = tk.Label(self.top)
        self.email_field_label.configure(anchor='w')
        self.email_field_label.configure(background="#d9d9d9")
        self.email_field_label.configure(compound='left')
        self.email_field_label.configure(disabledforeground="#a3a3a3")
        self.email_field_label.configure(foreground="#000000")
        self.email_field_label.configure(text='''Email''')

########## Show the last line of code I pasted.

        self.altname_field_label = tk.Label(self.top)
        self.altname_field_label.configure(anchor='w')
        self.altname_field_label.configure(background="#d9d9d9")
        self.altname_field_label.configure(compound='left')
        self.altname_field_label.configure(disabledforeground="#a3a3a3")
        self.altname_field_label.configure(foreground="#000000")
        self.altname_field_label.configure(text='''Alternate Name''')

        self.altphone_field_label = tk.Label(self.top)
        self.altphone_field_label.configure(anchor='w')
        self.altphone_field_label.configure(background="#d9d9d9")
        self.altphone_field_label.configure(compound='left')
        self.altphone_field_label.configure(disabledforeground="#a3a3a3")
        self.altphone_field_label.configure(foreground="#000000")
        self.altphone_field_label.configure(text='''Alt Phone''')

        self.altphone2_field_label = tk.Label(self.top)
        self.altphone2_field_label.configure(anchor='w')
        self.altphone2_field_label.configure(background="#d9d9d9")
        self.altphone2_field_label.configure(compound='left')
        self.altphone2_field_label.configure(disabledforeground="#a3a3a3")
        self.altphone2_field_label.configure(foreground="#000000")
        self.altphone2_field_label.configure(text='''Alt Phone 2''')

        self.altemail_field_label = tk.Label(self.top)
        self.altemail_field_label.configure(anchor='w')
        self.altemail_field_label.configure(background="#d9d9d9")
        self.altemail_field_label.configure(compound='left')
        self.altemail_field_label.configure(disabledforeground="#a3a3a3")
        self.altemail_field_label.configure(foreground="#000000")
        self.altemail_field_label.configure(text='''Alt Email''')

        self.model_field_label = tk.Label(self.top)
        self.model_field_label.configure(anchor='w')
        self.model_field_label.configure(background="#d9d9d9")
        self.model_field_label.configure(compound='left')
        self.model_field_label.configure(disabledforeground="#a3a3a3")
        self.model_field_label.configure(foreground="#000000")
        self.model_field_label.configure(text='''Model''')

########## Show only the last line I pasted.

        self.serial_field_label = tk.Label(self.top)
        self.serial_field_label.configure(anchor='w')
        self.serial_field_label.configure(background="#d9d9d9")
        self.serial_field_label.configure(compound='left')
        self.serial_field_label.configure(disabledforeground="#a3a3a3")
        self.serial_field_label.configure(foreground="#000000")
        self.serial_field_label.configure(text='''Serial''')

        self.callnum_label.place(relx=0.365, rely=0.050, height=20, width=135)
        self.dspnum_label.place(relx=0.500, rely=0.050, height=20, width=135)
        self.name_field_label.place(relx=0.365, rely=0.115, height=20, width=100)
        self.comp_field_label.place(relx=0.365, rely=0.180, height=20, width=100)
        self.addr1_field_label.place(relx=0.365, rely=0.245, height=20, width=100)
        self.addr2_field_label.place(relx=0.365, rely=0.310, height=20, width=100)
        self.city_field_label.place(relx=0.365, rely=0.375, height=20, width=100)
        self.state_field_label.place(relx=0.518, rely=0.375, height=20, width=50)
        self.zip_field_label.place(relx=0.555, rely=0.375, height=20, width=50)
        self.phone1_field_label.place(relx=0.365, rely=0.440, height=20, width=46)
        self.phone2_field_label.place(relx=0.435, rely=0.440, height=20, width=49)
        self.email_field_label.place(relx=0.505, rely=0.440, height=20, width=35)
        self.altname_field_label.place(relx=0.365, rely=0.505, height=20, width=94)
        self.altphone_field_label.place(relx=0.365, rely=0.570, height=20, width=64)
        self.altphone2_field_label.place(relx=0.435, rely=0.570, height=20, width=64)
        self.altemail_field_label.place(relx=0.505, rely=0.570, height=20, width=53)
        self.model_field_label.place(relx=0.365, rely=0.635, height=20, width=40)
        self.serial_field_label.place(relx=0.435, rely=0.635, height=20, width=34)

        self.menubar = tk.Menu(top, font="-family {Segoe UI} -size 9", bg=_bgcolor
                ,fg=_fgcolor)
        top.configure(menu = self.menubar)

########## Show the last line of code I pasted.

        self.sub_menu = tk.Menu(self.menubar, activebackground='beige'
                ,activeborderwidth=1, activeforeground='black'
                ,background='#d9d9d9', borderwidth=1, disabledforeground='#a3a3a3'
                ,font="-family {Segoe UI} -size 9", foreground='#000000'
                ,tearoff=0)
        self.menubar.add_cascade(compound='left', label='File', menu=self.sub_menu,)
        self.sub_menu.add_command(compound='left',label='Load')
        self.sub_menu.add_command(compound='left',label='Save')
        self.sub_menu.add_command(compound='left',label='Options')
        self.sub_menu.add_command(compound='left',label='Exit')
        self.sub_menu1 = tk.Menu(self.menubar, activebackground='beige'
                ,activeborderwidth=1, activeforeground='black'
                ,background='#d9d9d9', borderwidth=1, disabledforeground='#a3a3a3'
                ,font="-family {Segoe UI} -size 9", foreground='#000000'
                ,tearoff=0)
        self.menubar.add_cascade(compound='left', label='Action'
                ,menu=self.sub_menu1, )
        self.sub_menu1.add_command(compound='left',label='Paste Info')
        self.sub_menu1.add_separator()
        self.sub_menu1.add_command(compound='left',label='First Email')
        self.sub_menu1.add_command(compound='left',label='Second Email')
        self.sub_menu1.add_command(compound='left',label='Third Email')
        self.sub_menu1.add_command(compound='left',label='Final Email')

        self.sub_menu1.add_separator()
        self.sub_menu1.add_command(compound='left',label='Send Email')
        self.sub_menu12 = tk.Menu(self.menubar, activebackground='beige'
                ,activeborderwidth=1, activeforeground='black'
                ,background='#d9d9d9', borderwidth=1, disabledforeground='#a3a3a3'
                ,font="-family {Segoe UI} -size 9", foreground='#000000'
                ,tearoff=0)
        self.menubar.add_cascade(compound='left', label='Help'
                ,menu=self.sub_menu12, )
        self.sub_menu12.add_command(compound='left',label='About')
  
########## Show the last line of code I pasted.

output_list = []

def copy_to_preview(gui, var, line, preview_field):
    global output_list

    FieldsMap = {"name": gui.name.get(),
        "addr1": gui.addr1.get(),
        "addr2": gui.addr2.get(),
        "CSZ": gui.city.get(),
        "state": gui.state.get(),
        "zip": gui.zip.get(),
        "phone1": gui.phone.get(),
        "phone2": gui.phone2.get(),
        "email": gui.email.get(),
        "altcontact": gui.altcontact.get(),
        "altphone1": gui.altphone.get(),
        "altphone2": gui.altphone2.get(),
        "altemail": gui.altemail.get(),
        "model": gui.model.get(),
        "serial": gui.serial.get(),
        "repair_type": gui.repair_type.get(),
        "delay_type": gui.delay_type.get(),
        "start_time": gui.starttime.get(),
        "end_time": gui.endtime.get(),
        "parts": gui.parts.get()
        }
    
    if var.get():
        found = any(item['text'] == line for item in output_list)
        if not found:
            for key in FieldsMap:
                line = line.replace("{" + str(key) + "}", FieldsMap[key])
            output_list.append({'text': line})
    else:
        output_list.remove(line)

    preview_text = "\n".join(item['text'] for item in output_list)
    preview_field.delete("1.0", "end")
    preview_field.insert("1.0", preview_text)

########## Show the last line of code I pasted.

def popup1(self, event, *args, **kwargs):
    self.Popupmenu1 = self._extracted_from_popup3_2(event)

def popup2(self, event, *args, **kwargs):
    self.Popupmenu2 = self._extracted_from_popup3_2(event)

def popup3(self, event, *args, **kwargs):
    self.Popupmenu3 = self._extracted_from_popup3_2(event)

# TODO Rename this here and in `popup1`, `popup2` and `popup3`
def _extracted_from_popup3_2(self, event):
    result = tk.Menu(self.top, tearoff=0)
    result.configure(background=_bgcolor)
    result.configure(foreground=_fgcolor)
    result.configure(activebackground=_ana2color)
    result.configure(activeforeground='black')
    result.configure(font="-family {Segoe UI} -size 9")
    result.post(event.x_root, event.y_root)
    return result

# The following code is added to facilitate the Scrolled widgets you specified.
class AutoScroll(object):
    '''Configure the scrollbars for a widget.'''
    def __init__(self, master):
        #  Rozen. Added the try-except clauses so that this class
        #  could be used for scrolled entry widget for which vertical
        #  scrolling is not supported. 5/7/14.
        try:
            vsb = ttk.Scrollbar(master, orient='vertical', command=self.yview)
        except Exception:
            pass
        hsb = ttk.Scrollbar(master, orient='horizontal', command=self.xview)

########## Show the last line of code I pasted.

        try:
            self.configure(yscrollcommand=self._autoscroll(vsb))
        except Exception:
            pass

        self.configure(xscrollcommand=self._autoscroll(hsb))
        self.grid(column=0, row=0, sticky='nsew')
        try:
            vsb.grid(column=1, row=0, sticky='ns')
        except Exception:
            pass
        hsb.grid(column=0, row=1, sticky='ew')
        master.grid_columnconfigure(0, weight=1)
        master.grid_rowconfigure(0, weight=1)
        # Copy geometry methods of master  (taken from ScrolledText.py)
        methods = tk.Pack.__dict__.keys() | tk.Grid.__dict__.keys() \
                              | tk.Place.__dict__.keys()
        for meth in methods:
            if meth[0] != '_' and meth not in ('config', 'configure'):
                setattr(self, meth, getattr(master, meth))

    @staticmethod
    def _autoscroll(sbar):
        '''Hide and show scrollbar as needed.'''
        def wrapped(first, last):
            first, last = float(first), float(last)
            if first <= 0 and last >= 1:
                sbar.grid_remove()
            else:
                sbar.grid()
            sbar.set(first, last)
        return wrapped

    def __str__(self):
        return str(self.master)

########## Show the last line of code I pasted.

def _create_container(func):
    '''Creates a ttk Frame with a given master, and use this new frame to
    place the scrollbars and the widget.'''
    def wrapped(cls, master, **kw):
        container = ttk.Frame(master)
        container.bind('<Enter>', lambda e: _bound_to_mousewheel(e, container))
        container.bind('<Leave>', lambda e: _unbound_to_mousewheel(e, container))
        return func(cls, container, **kw)
    return wrapped

class ScrolledListBox(AutoScroll, tk.Listbox):
    '''A standard Tkinter Listbox widget with scrollbars that will
    automatically show/hide as needed.'''
    @_create_container
    def __init__(self, master, **kw):
        tk.Listbox.__init__(self, master, **kw)
        AutoScroll.__init__(self, master)
    def size_(self):
        return tk.Listbox.size(self)

import platform
def _bound_to_mousewheel(event, widget):
    child = widget.winfo_children()[0]
    if platform.system() in ['Windows', 'Darwin']:
        child.bind_all('<MouseWheel>', lambda e: _on_mousewheel(e, child))
        child.bind_all('<Shift-MouseWheel>', lambda e: _on_shiftmouse(e, child))
    else:
        child.bind_all('<Button-4>', lambda e: _on_mousewheel(e, child))
        child.bind_all('<Button-5>', lambda e: _on_mousewheel(e, child))
        child.bind_all('<Shift-Button-4>', lambda e: _on_shiftmouse(e, child))
        child.bind_all('<Shift-Button-5>', lambda e: _on_shiftmouse(e, child))

def _unbound_to_mousewheel(event, widget):
    if platform.system() in ['Windows', 'Darwin']:
        widget.unbind_all('<MouseWheel>')
        widget.unbind_all('<Shift-MouseWheel>')
    else:
        widget.unbind_all('<Button-4>')
        widget.unbind_all('<Button-5>')
        widget.unbind_all('<Shift-Button-4>')
        widget.unbind_all('<Shift-Button-5>')

########## Show the last line of code I pasted.

def _on_mousewheel(event, widget):
    if platform.system() == 'Windows':
        widget.yview_scroll(-1*int(event.delta/120),'units')
    elif platform.system() == 'Darwin':
        widget.yview_scroll(-1*int(event.delta),'units')
    elif event.num == 4:
        widget.yview_scroll(-1, 'units')
    elif event.num == 5:
        widget.yview_scroll(1, 'units')

def _on_shiftmouse(event, widget):
    if platform.system() == 'Windows':
        widget.xview_scroll(-1*int(event.delta/120), 'units')
    elif platform.system() == 'Darwin':
        widget.xview_scroll(-1*int(event.delta), 'units')
    elif event.num == 4:
        widget.xview_scroll(-1, 'units')
    elif event.num == 5:
        widget.xview_scroll(1, 'units')

def start_up():
        SvcNowParser_SUCCESS1.main()

if __name__ == '__main__':
    SvcNowParser_SUCCESS1.main()

########## This is ParserGUI.py. Do you have all the code I pasted?