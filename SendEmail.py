
# coding: utf-8

# In[2]:





# In[ ]:


import os
from tkinter import *
from tkinter import filedialog
from tkinter import messagebox
from tkinter import simpledialog
import tkinter as tk
import pandas as pd
import numpy as np
import socket
import smtplib
content = 'apple'
file_path = 'squarebot'
tmp=''
#FUNCTIONS
def Analyse(DB,E):
    global tmp
    DatabasePath= DB
    EventPath = E

    eventName = EventPath.split("/")[-1:][0].split(".")[0]
    mergeOn='Matric #'
    totalFrom='Name of Course'
    emailColumn = 'NUS Email'

    xl_fileDB = pd.ExcelFile(DatabasePath)
    DB = {sheet_name: xl_fileDB.parse(sheet_name) 
              for sheet_name in xl_fileDB.sheet_names}

    #Get the first Sheet Name
    DBsheet = xl_fileDB.sheet_names[0]

    xl_fileE = pd.ExcelFile(EventPath)
    Event = {sheet_name: xl_fileE.parse(sheet_name) 
              for sheet_name in xl_fileE.sheet_names}

    Eventsheet = xl_fileE.sheet_names[0]

    #sum from column index
    df = xl_fileDB.parse(xl_fileDB.sheet_names[0])


    #check if already imported
    colHeaders = list(df.columns.values)
    if (eventName in colHeaders):
        messagebox.showinfo ("Success","This event has already been IMPORTED")
        quit ()

    #Wrong file order
    if (totalFrom not in colHeaders):
        messagebox.showinfo ("Failure","Please check Input Files, Maybe in Wrong Order")
        quit ()

    index=0
    if totalFrom!="":
        index=df.columns.get_loc(totalFrom)

    Event[Eventsheet][eventName]=1
    s1 = pd.merge(DB[DBsheet], Event[Eventsheet], how='left', on=[mergeOn])
    s2 = Event[Eventsheet][~Event[Eventsheet][mergeOn].isin(s1[mergeOn])]
#     s2=s1[(~s1[mergeOn].isin(Event[DBsheet][mergeOn]))]
# s2 is People that is not from UTR
    writer1 = pd.ExcelWriter('./notFromUTR.xlsx')
    s2.to_excel(writer1,'Sheet1')
    writer1.save()

    if ('Total' in colHeaders):
        del s1['Total']

    s1['Total']=0
    s1['Total']=s1.iloc[:,index:].sum(axis=1)

    writer = pd.ExcelWriter('./output.xlsx')
    s1.to_excel(writer,'Sheet1')
    writer.save()

    if (len(list(s1[s1["Total"] == 0][emailColumn]))==0):
        messagebox.showinfo ("Success","All residents have participated")
        quit ()

    tmp = s1[s1["Total"] == 0][emailColumn]
    tmp[tmp!='[]']
    toBeEmailed = ";".join(str(x) for x in tmp)
#     toBeEmailed = ";".join(list(s1[s1["Total"] == 0][emailColumn]!=""))
    return toBeEmailed



def fetch(entries):
   for entry in entries:
      field = entry[0]
      text  = entry[1].get()
      print('%s: "%s"' % (field, text)) 

def makeform(root, fields):
   entries = []
   for field in fields:
      row = Frame(root)
      lab = Label(row, width=20, text=field, anchor='w')
      ent = Entry(row)
      row.pack(side=TOP, fill=X, padx=5, pady=5)
      lab.pack(side=LEFT)
      ent.pack(expand=NO, fill=X, padx = 5)
      entries.append((field, ent))
   return entries

def showMessage(s,s1):
    messagebox.showinfo(s, s1)


def browsefunc(): #browse button to search for files
    filename = filedialog.askopenfilename(filetypes=(("Excel", "*.xlsx"),
                                           ("All files", "*.*") ))
    # infile = open(filename, 'r')
    # content = infile.read()
    #pathadd = os.path.dirname(filename)+filename
    pathadd = filename
    file_path1.set(pathadd)
    return content

def browsefunc2(): #browse button to search for files
    filename2 = filedialog.askopenfilename(filetypes=(("Excel", "*.xlsx"),
                                           ("All files", "*.*") ))
    # infile = open(filename2, 'r')
    # content = infile.read()
    #pathadd = os.path.dirname(filename2)+filename2
    pathadd = filename2
    file_path2.set(pathadd)
    return content

def browsefunc3(): #browse button to search for files
    directory = filedialog.askdirectory(initialdir='.')
    directoryname.set(directory)
    return content

def process_file(): #process reconciliation code
#     print('------------------------------')
#     print(file_path1.get())
#     print(file_path2.get())
    
    try:
        displayText = Analyse(file_path1.get(),file_path2.get())
        CustomDialog(root, title="List of Emails to be sent", text=displayText)
    except (FileNotFoundError,UnboundLocalError) as e:
        showMessage("Error!", "Certain File not Found")
        print (e)
        return ""
    except (KeyError) as e:
        showMessage("Error!","Remember to add header for event file. E.g: Matric #")
        print (e)
        return ""
    return displayText
    
def process_file1(): #process reconciliation code
#     print('------------------------------')
#     print(file_path1.get())
#     print(file_path2.get())
    try:
        displayText = Analyse(file_path1.get(),file_path2.get())
    except (FileNotFoundError):
        showMessage("Error!", "File NOT Found")
        return ""
    return displayText

def sendEmail(s):
    print('------------------------------')
    displayText=process_file1()
    emailList =displayText.split(';')
    print (emailAdd.get())
    print (emailPass.get())
    print (emailSub.get())
    print (s.get("1.0",'end-1c'))
      
    
    email = str(emailAdd.get())
    password=str(emailPass.get())
    subject = str(emailSub.get())
    content = str(s.get("1.0",'end-1c'))

    SERVER = "smtp.office365.com"
    FROM = email
    TO = emailList # must be a list

    SUBJECT = subject
    TEXT = content 

    # Prepare actual message
    message = """"From: %s\r\nTo: %s\r\nSubject: %s\r\n
    %s
    """ % (FROM, ", ".join(TO), SUBJECT, TEXT)

    # Send the mail
    import smtplib
    try:
        server = smtplib.SMTP(SERVER,587)
        # server.connect(SERVER,25)

        server.ehlo()
        server.starttls()
        server.ehlo()
        try:
            server.login(email, password)
            server.sendmail(FROM, TO, message)
            showMessage ("Successful","Successfully sent email")
            CustomDialog(root, title="Successfully Sent to these", text=displayText)
        finally:
            server.quit()
    except (smtplib.SMTPRecipientsRefused) as e:
        showMessage ("Email sent Unsuccessful","Please Check Recipient's Email")
        print ("Fail to send email")
        print("Please Check Recipient's Email")
        refused = e.recipients
        showMessage ("Email sent Unsuccessful",refused)
    except (smtplib.SMTPAuthenticationError, smtplib.SMTPException) as e:
        showMessage ("Unsuccessful","Check your email address and Password")
        print ("Fail to send email")
        print ("Check your email address and Password")
    except (socket.gaierror) as e:
        showMessage("No Internet","Please Check that you are online")
    except (TypeError) as e:
        showMessage ("Error!","Please insert Email and Password")
    # 
class CustomDialog(simpledialog.Dialog):

    def __init__(self, parent, title=None, text=None):
        self.data = text
        simpledialog.Dialog.__init__(self, parent, title=title)

    def body(self, parent):

        self.text = tk.Text(self, width=40, height=4)
        self.text.pack(fill="both", expand=True)

        self.text.insert("1.0", self.data)

        return self.text



#GUI

root = Tk()

root.title('UTR Emailing App')
root.geometry("698x350")

mf = Frame(root)
mf.pack()

f1 = Frame(mf, width=700, height=500) #file1
f1.pack(fill=X)
f2 = Frame(mf, width=700, height=500) #file2
f2.pack(fill=X)
f4 = Frame(mf, width=700, height=500) #reconcile button
f4.pack(fill=X)
f5 = Frame(mf, width=700, height=500)
f5.pack(fill=X)

#email Pass
f6 = Frame(mf,width=700, height=500)
f6.pack(fill=X)

#Email Subject
f7 = Frame(mf, width=700, height=500)
f7.pack(fill=X)

#Email Content
f8 = Frame(mf, width=700, height=500)
f8.pack(fill=X)

f9 = Frame(mf, width=700, height=500)
f9.pack(fill=X)

file_path1 = StringVar()
file_path2 = StringVar()
directoryname = StringVar()

Label(f1,text="Main xlsx File ").grid(row=0, column=0, sticky='e') #file1 button
entry1 = Entry(f1, width=50, textvariable=file_path1)
entry1.grid(row=0,column=1,padx=2,pady=2,sticky='we',columnspan=25)

Label(f2,text="Event xlsx File").grid(row=0, column=0, sticky='e') #file2 button
entry2 = Entry(f2, width=50, textvariable=file_path2)
entry2.grid(row=0,column=1,padx=2,pady=2,sticky='we',columnspan=25)

Button(f1, text="Browse", command=browsefunc).grid(row=0, column=27, sticky='ew', padx=8, pady=4)#file1 button
Button(f2, text="Browse", command=browsefunc2).grid(row=0, column=27, sticky='ew', padx=8, pady=4)#file2 button
# Button(f3, text="Browse", command=browsefunc3).grid(row=0, column=27, sticky='ew', padx=8, pady=4)#destination folder button
# Button(f3, text="GetValue", width=32, command=(lambda e=ents: fetch(e))).grid(sticky='ew', padx=10, pady=10)#reconcile button

Button(f4, text="Analyse", width=32, command=lambda: process_file()).grid(sticky='ew', padx=10, pady=10)#reconcile button


emailAdd = StringVar()
emailPass = StringVar()
emailSub = StringVar()
# emailContent = StringVar()

Label(f5,text="Email Address  ").grid(row=0, column=0, sticky='e')
entry3 = Entry(f5, width=50, textvariable=emailAdd)
entry3.grid(row=0,column=1,padx=2,pady=2,sticky='we',columnspan=25)


Label(f6,text="Email Password ").grid(row=0, column=0, sticky='e')
entry4 = Entry(f6, width=50, textvariable=emailPass,show='*')
entry4.grid(row=0,column=1,padx=2,pady=2,sticky='we',columnspan=25)


Label(f7,text="Subject        ").grid(row=0, column=0, sticky='e')
entry5 = Entry(f7, width=50, textvariable=emailSub)
entry5.grid(row=0,column=1,padx=2,pady=2,sticky='we',columnspan=25)

Label(f8,text="Content        ").grid(row=0, column=0, sticky='e')
# entry6 = Entry(f8, width=50, textvariable=emailContent)
# entry6.grid(row=0,column=1,padx=2,pady=2,sticky='we',columnspan=25)
emailContent = Text(f8, width = 50, height = 10, takefocus=0)
emailContent.grid(row=0, column=1,sticky='e', padx=2, pady=2)


Button(f9, text="Send Email", width=32, command=lambda: sendEmail(emailContent)).grid(sticky='ew', padx=10, pady=10)#reconcile button



root.mainloop()


# In[23]:





# In[26]:





# In[17]:




