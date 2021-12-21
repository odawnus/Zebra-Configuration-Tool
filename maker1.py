##========================================
##Created by : Collin Preston
##Created date: 8/31/2020
##========================================

"""
Zebra Configuration Tool

Tool help with configuring zebra printers to connect to a network.
Designed program so Tech make configuration files for devices ahead of time
and reduces chances that a device will be incorrectly configured. For program to work
please install the full zebra driver package.
"""

##==========================
## Imports
##==========================
import tkinter as tk
from tkinter import *
import tkinter.font as tkFont
from tkinter import messagebox
from tkinter import filedialog 
import os
import win32print
from operator import itemgetter



#Fuction determines which radio button is selected, this value will be
#used by the submitValues function to create correct file type.
def sel(var):
    selection = var.get()
    return selection
                       
def img_resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller """
    try:
        # PyInstaller creates a temp folder and stores path in _MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)


#Use win32print to show all print devices on the installed on the computer and stores them in a variable as a list
def onlineprinters():
    #When emnumerating the list it returns a value(atrribute), if that value is equal to 0 then that device is currently online.
    printers = win32print.EnumPrinters(win32print.PRINTER_ENUM_LOCAL, None)
    p1 = list(map(itemgetter(2), printers))
    online = []
    for i in p1:
        handle = win32print.OpenPrinter(i)
        attributes = win32print.GetPrinter(handle)[13]
        pop = attributes & 0x00000400   
        if pop == 0:
            online.append(i)
        #All Zebra printers by default start with Z, this removes all printers that do not start with Z    
        start_letter = 'Z'
        with_z = list(filter(lambda x: x.startswith(start_letter), online))
    return with_z







#GUI Window
window = Tk()
window.geometry("650x200") 
window.title("Zebra Configuration Tool")

#inputs from the GUI are sent to zebra() function
fontstyle1 = tkFont.Font(family="Lucida Grande", size=20)
fontstyle2 = tkFont.Font(family="Lucida Grande", size=10)
fontstyle3 = tkFont.Font(family="Lucida Grande", size=8)
fontstyle4 = tkFont.Font(family="Showcard Gothic", size=20)
fontstyle5 = tkFont.Font(family="Lucida Grande", size=7)
window.iconbitmap(default=img_resource_path('piedmonticon.ico'))
Lb1 = Listbox(window, height=6,width=30,bg="white",bd=3)
Lb1.place(x=440, y=30)
Lb1label= Label(window,font=fontstyle2, text="Select Printer")
Lb1label.place(x=490,y=5)

def populatebox():
    Lb1.delete(0, END)
    for i in onlineprinters():
        Lb1.insert("end", i)

#Below is what you see visually, GUI
btn = Button(window, text="Update list", command = lambda: populatebox())
btn.place(x=440, y=140)
btn2 = Button(window, text="Send File to Printer", command = lambda: printValues())
btn2.place(x=520, y=140)
btn3 = Button(window, text="Print Info", command = lambda: printinfo())
btn3.place(x=440, y=170)
# Header
header = Label(window,font=fontstyle1,  text="Zebra Configuration Tool" )
header.grid(row=1, column=2)
# Entry Boxes and Labels
l1 = Label(window,font=fontstyle2, text="Printer Name")
l1.grid(row=3,column=1)
l1.focus()
e1 = Entry(window,show=None, bd =3, width=50)
e1.grid(row=3,column=2)
l2 = Label(window,font=fontstyle2, text="IP Address")
l2.grid(row=4,column=1)
e2 = Entry(window,show=None, bd =3, width=50,)
e2.grid(row=4,column=2)
l3 = Label(window,font=fontstyle2, text="Subnet Mask",)
l3.grid(row=5,column=1)
e3 = Entry(window,show=None, bd =3, width=50,)
e3.grid(row=5,column=2)
l4 = Label(window,font=fontstyle2, text="Gateway")
l4.grid(row=6,column=1)
e4 = Entry(window,show=None, bd =3, width=50,)
e4.grid(row=6,column=2)      
#Radio buttons for wired/wireless selection
var = IntVar()
R1 = Radiobutton(window, text="ZT230/GK420", variable=var, value=1,command=lambda: sel(var))
R1.place(x=240,y=130)
R2 = Radiobutton(window, text="QLN320/ZQ620", variable=var, value=2,command=lambda: sel(var))
R2.place(x=240,y=150)
infoboxcheck = 1
infobtn = Button(window, width=6,height=1,font=fontstyle3,text="?",bg="firebrick4",fg="white",activebackground="white",activeforeground="firebrick4", command = lambda: about())
infobtn.place(x=10,y=160) 

#Info Button
def about3():
    global infoboxcheck
    infoboxcheck = 1

def about2(popup):
    popup.destroy()


def about():
    global infoboxcheck
    if infoboxcheck == 1:
        infoboxcheck = 0
        root_x = window.winfo_rootx()
        root_y = window.winfo_rooty()
        popup = Tk()
        win_x = root_x + 150
        win_y = root_y + 200
        popup.geometry(f'+{win_x}+{win_y}')
        popup.wm_title("Program Info")
        
        popup.configure(background='sky blue')
        infofont1 = tkFont.Font(family="Lucida Grande", size=20, weight=tkFont.BOLD)        
        infolabel = Label(popup, font = infofont1, text="Zebra Configuration Tool\nVersion 2.0\nCopyright 2021 C.Preston",bg="sky blue",fg="black")
        infolabel.pack(side="top", fill="x", pady=10)
        infolabel2 = Label(popup, text="Program For Use Only By Those With Permission\nSend Questions/Comments/Bug Reports to\nCollin.Preston1@piedmont.org",bg="sky blue", font=fontstyle4)
        infolabel2.pack( fill="x", pady=10)
        popup.attributes("-topmost", True)
        popup.overrideredirect(True)
        B1 = Button(popup, text="Close",bg="azure", command = lambda: [about2(popup),about3()])
        B1.pack(side="bottom")
    else:
        pass




def submitValues():
#content for file for wired devices
    if sel(var) == 1:
        content = str(
"^XA\n"
"^ND2,P,{},{},{}\n"
"^NBC\n"
"^NC1\n"
"^NPP\n"
"^NN{}\n"
"^PW406\n"
"^XZ\n"
"^XA\n"
"~SD15\n"
"~TA000\n"
"~JSN\n"
"^SZ2\n"
"^LL253\n"
"^PON\n"
"^PR5,5\n"
"^PMN\n"
"^MNY\n"
"^LS0\n"
"^MTD\n"
"^MMT,N\n"
"^MPE\n"
"^MPP\n"
"^MPM\n"
"^XZ\n"
"^XA\n"
"^JUS\n"
"^XZ\n".format(e2.get(),e3.get(),e4.get(),e1.get()))
        
    elif sel(var) == 2:
        #content for file for wireless devices
        content = str(
'^XA\n'
'^WIP,{},{},{}\n'
'^NC2\n'
'^NPP\n'
'^KC0,0,,\n'
'^NN{}\n'
'^WAD,D\n'
'^WEOFF,1,O,H,,,,\n'
'^WP0,0\n'
'^WR,,,,100\n'
'^WSPHCD,I,L,,,\n'
'^PW406\n'
'~SD15\n'
'~TA000\n'
'~JSN\n'
'^SZ2\n'
'^LL253\n'
'^NBS\n'
'^WLOFF,,\n'
'^WKOFF,,,,\n'
'^WX09,000000000000000000000000000000000000000000000000000000\n'
'^PON\n'
'^PR5,5\n'
'^PMN\n'
'^MNY\n'
'^LS0\n'
'^MTD\n'
'^MMT,N\n'
'^MPE\n'
'! U1 setvar "wlan.international_mode" "off"\n'
'! U1 setvar "wlan.allowed_band" "all"\n'
'! U1 SPEED 5\n'
'! U1 setvar "print.tone" "0"\n'
'! U1 setvar "media.type" "label"\n'
'! U1 setvar "media.sense_mode" "gap"\n'
'! U1 setvar "media.darkness_mode" "15"\n'
'!  U1 setvar "device.prompted_network_reset" "yes"\n'
'~JC^JUS^XZ\n'.format(e2.get(),e3.get(),e4.get(),e1.get()
))
    if sel(var) == 0:
        failure = messagebox.showinfo("Error","Please Select Correct Model Type")
        return
    if e1.get() =="":
        failure = messagebox.showinfo("Error","Please Enter Name for Device")
        return
    elif e2.get() == "":
        failure = messagebox.showinfo("Error","Please Enter IP Address")
        return
    elif e3.get() == "":
        failure = messagebox.showinfo("Error","Please Enter Subnet Mask")
        return
    elif e4.get() == "":
        failure = messagebox.showinfo("Error","Please Enter Gateway")
        return
    f = filedialog.asksaveasfile(mode='w',initialfile = e1.get(),defaultextension=".zpl", filetypes =[("Zebra ZPL",".zpl")])
    if f is None:
        return
    f.write(content)
    f.close() 
    return


#Using the list of currently install devices and showing them in a GUI Window.
#Will not let you proceed if no device is selected. Will prompt error.
#Send file to printer button on GUI
#Send jobs to printer once a file is selected, printer configs and firmware can be sent.
def printValues():
    items = list(Lb1.curselection())
    if(len(items) == 0):
        failure = messagebox.showinfo("Error","Please Select Printer")
    else:    
        window.filename =  filedialog.askopenfilename(initialdir = "/áéá",title = "Select file",filetypes = (("Zebra ZPL","*.zpl"),("all files","*.*")))
        if window.filename == "":
            return
        file = open(window.filename)
        output = file.read()
        p = win32print.OpenPrinter(Lb1.get(Lb1.curselection()))
        win32print.StartDocPrinter(p, 1, (output, None, 'RAW'))
        win32print.WritePrinter(p, (output.encode()))
        win32print.EndDocPrinter(p)
        file.close()
        return

#Using the list of currently install devices and showing them in a GUI Window.
#Will not let you proceed if no device is selected. Will prompt error.
#Printer Info button on GUI
#Sends a job to print, Printer information is printed   
def printinfo():
    items = list(Lb1.curselection())
    if(len(items) == 0):
        failure = messagebox.showinfo("Error","Please Select Printer")
    else:    
        output = "~WC"
        p = win32print.OpenPrinter(Lb1.get(Lb1.curselection()))
        win32print.StartDocPrinter(p, 1, (output, None, 'RAW'))
        win32print.WritePrinter(p, (output.encode()))
        win32print.EndDocPrinter(p)
        return


#Save to file
submit = tk.Button(window, text="Save to File",command=lambda: submitValues()) 
submit.place(x=160,y=140)


window.mainloop()
