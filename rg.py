from Tkinter import *
from win32com.client import Dispatch, constants
import win32gui, win32con, win32com.client
import datetime
import tkMessageBox

def initialize1():
    root = Tk()
    root.title("Status Rack Report")
    root.iconbitmap('if_letter_S_red_1553075.ico')

    L02 = Label(root, text="Manager")
    L02.grid(row=0, column=0, sticky= W)
    variable1 = StringVar(root)
    variable1.set("Doug Fong") # default value
    w02 = OptionMenu(root, variable1,"Doug Fong", "Willis Wong",\
                     "Stuart Robson", "Ben Mansell",\
                     "Paul Laca", "David Nguyen",\
                     "George Binns", "Nasser Iranikhah")
    w02.grid(row=0, column=1, sticky= W)

    LCC = Label(root, text="CC")
    LCC.grid(row=1, column =0, sticky= W)
    ECC = Entry(root)
    ECC.grid(row=1, column=1, sticky= W)

    L03 = Label(root, text="Customer Name")
    L03.grid(row=2, column =0, sticky= W)
    E0 = Entry(root)
    E0.grid(row=2, column=1, sticky= W)

    L0 = Label(root, text="Rack Name (*One rack per report)")
    L0.grid(row=3, column =0, sticky= W)
    T0 = Entry(root)
    T0.grid(row=3, column=1, sticky= W)

    L01 = Label(root, text="ATP Container#")
    L01.grid(row=4, column =0, sticky= W)
    T01 = Entry(root)
    T01.grid(row=4, column=1, sticky= W)

    L1 = Label(root, text="Key Activities Completed on Day")
    L1.grid(row=5, column =0, sticky= W)
    T1 = Text(root, height=5, width=75)
    T1.delete("1.0",END)
##    T1.insert("1.0","a.\nb.\nc.")
    T1.grid(row=6, column=1, sticky= W)
    S1 = Scrollbar(root)
    S1.grid(row=6,column=2,sticky=NS)
    S1.config(command=T1.yview)
    T1.config(yscrollcommand=S1.set)

    L2 = Label(root, text="ATP_Status (current stage of rack ATP)")
    L2.grid(row=7, column=0, sticky= W)
    variable = StringVar(root)
    variable.set("Please pick up a status") # default value
    w = OptionMenu(root, variable, \
                   "Pre-Integration", "Integration", "Dry Run", "ATP")
    w.grid(row=7, column=1, sticky= W)

    L3 = Label(root, text="Justification for software updates")
    L3.grid(row=8, column=0, sticky= W)
    T3 = Text(root, height=5, width=75)
    T3.delete("1.0",END)
##    T3.insert("1.0","a.\nb.\nc.")
    T3.grid(row=9, column=1, sticky= E)
    S3 = Scrollbar(root)
    S3.grid(row=9,column=2,sticky=NS)
    S3.config(command=T3.yview)
    T3.config(yscrollcommand=S3.set)

    L4 = Label(root, text="Critical issues Raised and potential impact")
    L4.grid(row=10, column=0, sticky= W)
    T4 = Text(root, height=5, width=75)
    T4.delete("1.0",END)
##    T4.insert("1.0","a.\nb.\nc.")
    T4.grid(row=11, column=1, sticky= W)
    S4 = Scrollbar(root)
    S4.grid(row=11,column=2,sticky=NS)
    S4.config(command=T4.yview)
    T4.config(yscrollcommand=S4.set)

    L5 = Label(root, text="Support Request")
    L5.grid(row=12, column=0, sticky= W)
    T5 = Text(root, height=5, width=75)
    T5.delete("1.0",END)
##    T5.insert("1.0","a.\nb.\nc.")
    T5.grid(row=13, column=1, sticky= W)
    S5 = Scrollbar(root)
    S5.grid(row=13,column=2,sticky=NS)
    S5.config(command=T5.yview)
    T5.config(yscrollcommand=S5.set)

    B1 = Button(root, text="Send!", command=lambda: sendText(variable1,ECC,E0,T0,T01,\
                                                             variable,T1,T3,T4,T5),\
                height = 5, width = 20)
    B1.grid(row=14,column=0,sticky=E)
    B2 = Button(root, text="Clear All", command=lambda: rmText(ECC,E0,T0,T01,T1,T3,T4,T5),\
                height = 5, width = 20)
##    B2 = Button(root, text="Clear All", command=lambda: orgText(T1))
    B2.grid(row=14,column=1)
    

    root.mainloop()

def rmText(ac,a1,a,b,c,d,e,f):
    result = tkMessageBox.askquestion("Warning",\
                                      "Are you sure you want to clear this form?",\
                                      icon='warning')
    if result == 'yes':
        ac.delete(0,END)
        a1.delete(0,END)
        a.delete(0,END)
        b.delete(0,END)
    ##    b.insert("1.0","a.\nb.\nc.")
        c.delete(1.0,END)
    ##    c.insert("1.0","a.\nb.\nc.")
        d.delete(1.0,END)
    ##    d.insert("1.0","a.\nb.\nc.")
        e.delete(1.0,END)
    ##    e.insert("1.0","a.\nb.\nc.")
        f.delete(1.0,END)

def sendText(mana,cc,cust,rack,atp,drop,key,just,criti,supp):
    cc_name = cc.get()
    cust_name = cust.get()
    rack_name = rack.get()
    atp_no = atp.get()
##    key_act = key.get("1.0",END)
##    justification = just.get("1.0",END)
##    critical = criti.get("1.0",END)
##    support = supp.get("1.0",END)

    now=datetime.date.today().strftime("%m/%d/%y")

    cc_name = cc_name.rstrip()
    cc_name = cc_name.strip()
    cust_name = cust_name.rstrip()
    cust_name = cust_name.strip()
    rack_name = rack_name.rstrip()
    rack_name = rack_name.strip()
    atp_no = atp_no.rstrip()
    atp_no = atp_no.strip()
##    rack_name = rack_name.replace(" ","")
##    atp_no = atp_no.replace(" ","")
    
    const=win32com.client.constants
    olMailItem = 0x0   
    obj = win32com.client.Dispatch("Outlook.Application")
    newMail = obj.CreateItem(olMailItem)
    newMail.Subject = "Rack Report -- "+ cust_name + "/" +rack_name +" -- "+now
    
##    newMail.Body = "The following represents the status report for "+rack_name+\
##                   ", ATP#"+atp_no+"\n\n"+\
##                   "(1)Key Activities Completed on Day\n"+orgText(key)+"\n"+\
##                   "(2)ATP_Status (current stage of rack ATP)\n"+drop.get()+"\n\n"+\
##                   "(3)Justification for Software updates\n"+orgText(just)+"\n"+\
##                   "(4)Critical issues Raised and potential Impact\n"+orgText(criti)+"\n"\
##                   "(5)Support Requested\n"+orgText(supp)
##    newMail.BodyFormat = 3
    newMail.HTMLBody = "The following represents the status report for "+rack_name+\
                   ", ATP#"+atp_no+"<br><br>"+\
                   "(1) Key Activities Completed on Day<br>"+orgText(key)+"<br>"+\
                   "(2) ATP_Status (current stage of rack ATP)<br>&nbsp;&nbsp;&nbsp;&nbsp;"+drop.get()+"<br><br>"+\
                   "(3) Justification for Software updates<br>"+orgText(just)+"<br>"+\
                   "(4) Critical issues Raised and potential Impact<br>"+orgText(criti)+"<br>"\
                   "(5) Support Requested<br>"+orgText(supp)
    newMail.To = toEmail(mana.get())
    newMail.Cc = cc_name
    newMail.Send()
    tkMessageBox.showinfo("sent","Status Report successfully submitted to "+mana.get()+"!")

def orgText(T):
    T=T.get("1.0",END)
    T_split=T.splitlines()
    result = ""
    index = 1
    for a in T_split:
        if a:
            result = result + "&nbsp;&nbsp;&nbsp;&nbsp;" + str(index) + ". " + a + "<br>"
            index = index + 1
##    print(len(T_split)) 
##    print(T_split)
##    print(result)
    return result
    
def toEmail(m_name):
    if m_name == "Doug Fong":
        result = "doug.fong@panasonic.aero"
    elif m_name == "Willis Wong":
        result = "willis.wong@panasonic.aero"
    elif m_name == "Stuart Robson":
        result = "stuart.robson@panasonic.aero"
    elif m_name == "Ben Mansell":
        result = "ben.mansell@panasonic.aero"
    elif m_name == "Paul Laca":
        result = "paul.laca@panasonic.aero"
    elif m_name == "David Nguyen":
        result = "david.nguyen@panasonic.aero"
    elif m_name == "George Binns":
        result = "george.binns@panasonic.aero"
    elif m_name == "Nasser Iranikhah":
        result = "nasser.iranikhah@panasonic.aero"
    return result

try:
    initialize1()
except Exception as e:
    tkMessageBox.showinfo("Error", str(e))

