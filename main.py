#from openpyxl import load_workbook
#from openpyxl import Workbook
from pyxll import xl_menu, create_ctp, CTPDockPositionFloating
from enum import global_flag_repr
from functools import partial

import tkinter as tk                                    # Tkinter's Tk class
import tkinter.ttk as ttk                               # Tkinter's Tkk class
import datetime

from tkinter import filedialog as fd
from PIL import ImageTk, Image
from tkinter import messagebox
from random import shuffle
from tkcalender import DateEntry

GUI = tk.Tk()
GUI.title("LAL Measurement")
GUI.geometry("1000x700")                                # Set the geometry of Tkinter frame
GUI.configure(bg = 'white')                             # Set background color
GUI.option_add("*Font", "Helvetica 12 bold")            # set the font and size for entire gui
GUI.option_add("*fg", "black")                          # set the text color, hex works too "#FFFFFF"
GUI.option_add("*bg", "white")                          # set the 

def resize_image(event):
    new_width = event.width
    new_height = event.height
    background_image = copy_of_image.resize((new_width, new_height))
    bkgnd_img = ImageTk.PhotoImage(background_image)
    lbl_photo.config(image = bkgnd_img)
    lbl_photo.background_image = bkgnd_img #avoid garbage collection

background_image = Image.open(r"C:\Users\Susan\OneDrive\Documents\Python\GUI_Excel\LAL.png")
copy_of_image = background_image.copy()
bkgnd_img = ImageTk.PhotoImage(background_image)

lbl_photo = ttk.Label(GUI, image = bkgnd_img)
lbl_photo.bind('<Configure>', resize_image)
lbl_photo.pack(fill=tk.BOTH, expand = True)

# Python is serial, so each widget will output in order placed below;

################################################         TODAYS DATE (AUTOFILL)           ################################################   
date = datetime.datetime.now()

lbl_cmd_date = tk.Label(                                # set the constant output text command to the user with instructions
    text="Today's Date is:",                            # set text for the operator to read
    bg = "white",                                       # set the background color, hex works too "#FFFFFF"
    width = 12,                                         # set the width of text box, measured in text units '0'. 50 = 50 zeros wide
    height= 1                                           # if height is 1, then no need to call it out.
) 
lbl_cmd_date.place(x=50,y=50)  

# %B = Month spelled out, %d = day of the month DD, %Y = Year YYYY. Longest month spelling is September, this width accounts for that.
lbl_out_date = tk.Label(GUI, locale='en_US', date_pattern='mm/dd/y').place(x=300, y=50)

################################################            OPERATOR INPUT             ################################################  

# 4 Labels to COMMAND the user what to type in the entry field
lbl_command_cred    = tk.Label(text="Enter Operator Credentials:",  bg= "white", width= 20).place(x=50,y=100)
lbl_command_WO      = tk.Label(text="Enter Work Order Number:",     bg= "white", width= 20).place(x=50,y=150)  
lbl_command_samp    = tk.Label(text="Enter Sample Size:",           bg= "white", width= 14).place(x=50,y=200) 
lbl_command_meas    = tk.Label(text="Select Measurement Size:",     bg= "white", width= 20).place(x=50,y=250)  

# 3 Entry fields to accept User Input
entry_cred  = tk.Entry(GUI, bg= "white", width= 10)
entry_cred.focus_set()                                  # This set's the cursor in the entry box to star typing in first.
entry_cred.place(x=300,y=100)
entry_WO    = tk.Entry(GUI, bg= "white", width= 10).place(x=300,y=150)
entry_samp  = tk.Entry(GUI, bg= "white", width= 10).place(x=300,y=200) 

# 4 Labels to DISPLAY / Re-iterate what type of data is going to be displayed
lbl_disp_cred   = tk.Label(GUI, text="Credentials:",        bg= "white", width= 9).place( x=45, y=450)
lbl_disp_WO     = tk.Label(GUI, text="Work Order Number:",  bg= "white", width= 16).place(x=50, y=500)
lbl_disp_samp   = tk.Label(GUI, text="Sample Size:",        bg= "white", width= 10).place(x=50, y=550)
lbl_disp_meas   = tk.Label(GUI, text="Measurement Size:",   bg= "white", width= 15).place(x=50, y=600)

# OUTPUT operator entry back for them to read and confirm the information is correct. This happens after pressing "Confirm".
lbl_out_cred    = tk.Label(GUI, text = "", bg= "white", width= 3).place( x=300, y=450) 
lbl_out_WO      = tk.Label(GUI, text= "",  bg= "white", width= 10).place(x=300, y=500)
lbl_out_samp    = tk.Label(GUI, text="",   bg= "white", width= 3) .place(x=300, y=550)

def display_cred(): # Display Credentials
   global entry
   cred = entry_cred.get()[:3]  # limit 3 characters
   lbl_out_cred.configure(text = cred)
   print(entry_cred.get()[:3])  # entry_cred is the variable we are passing, limit 3 characters print can be removed after R&D

def display_WO():                        
   global entry
   WO = entry_WO.get()[:10]
   lbl_out_WO.configure(text = WO)
   print(entry_WO.get()[:10])

def display_samp():                        
   global entry
   samp = entry_samp.get()[:2]
   lbl_out_samp.configure(text = samp)
   print(entry_samp.get()[:2])

def display_015_040(text):
   disp_meas = tk.Entry(GUI, width= 3)
   disp_meas.insert(0,text)
   disp_meas.place(x=300, y=600)
   print(text)

btn_cred = tk.Button(GUI, text="Confirm", bg= "grey", width= 10, command = display_cred).place(x=450,y=90)
btn_WO   = tk.Button(GUI, text="Confirm", bg= "grey", width= 10, command = display_WO).place(x=450,y=140)  
btn_samp = tk.Button(GUI, text="Confirm", bg= "grey", width= 10, command = display_samp).place(x=450,y=190)  
b015     = tk.Button(GUI, text="015",     bg= "grey", width= 5, command  = partial(display_015_040,"015")).place(x=300,y=250) 
b040     = tk.Button(GUI, text="040",     bg= "grey", width= 5, command  = partial(display_015_040,"040")).place(x=400,y=250)

################################################             EXIT BUTTON             ################################################   
def exit_application():
    msg_box = tk.messagebox.askquestion('Exit', 'Are you sure you want to exit the application?', icon='warning')
    if msg_box == 'yes':
        GUI.destroy()
    else:
        tk.messagebox.showinfo('Exit', 'Thanks for staying, please continue.')

but_exit = tk.Button(GUI, text="Exit", bg = "grey", width= 5, command=exit_application).place(x=900,y=630)
# Must be at the end of the program in order for the application to run b/c windows is constantly updating
GUI.mainloop() 