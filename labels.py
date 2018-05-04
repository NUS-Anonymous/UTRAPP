import os
from tkinter import *
from tkinter import filedialog
from tkinter import messagebox
from tkinter import simpledialog
import tkinter as tk
import pandas as pd
import numpy as np

content = 'apple'
file_path = 'squarebot'

#FUNCTIONS


def Analyse(DB,E):
    DatabasePath= DB
    EventPath = E

    eventName = EventPath.split("/")[-1:][0].split(".")[0]
    mergeOn='Name'
    totalFrom='Loc'

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

    if ('Total' in colHeaders):
        del s1['Total']

    s1['Total']=0
    s1['Total']=s1.iloc[:,index:].sum(axis=1)

    writer = pd.ExcelWriter('./output.xlsx')
    s1.to_excel(writer,'Sheet1')
    writer.save()

    if (len(list(s1[s1["Total"] == 0]['Emails']))==0):
        messagebox.showinfo ("Success","All residents have participated")
        quit ()


    toBeEmailed = ";".join(list(s1[s1["Total"] == 0]['Emails']))
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

def process_file(content): #process reconciliation code
    # print('------------------------------')
    # print(file_path1.get())
    # print(file_path2.get())
    # messagebox.showinfo("List of Emails", "Hello World")
    displayText = Analyse(file_path1.get(),file_path2.get())
    CustomDialog(root, title="List of Emails to be sent", text=displayText)

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

root.title('Reconciliation Converter')
root.geometry("698x150")

mf = Frame(root)
mf.pack()

f1 = Frame(mf, width=600, height=250) #file1
f1.pack(fill=X)
f2 = Frame(mf, width=600, height=250) #file2
f2.pack(fill=X)
f4 = Frame(mf, width=600, height=250) #reconcile button
f4.pack(fill=X)

file_path1 = StringVar()
file_path2 = StringVar()
directoryname = StringVar()

Label(f1,text="Main Excel File").grid(row=0, column=0, sticky='e') #file1 button
entry1 = Entry(f1, width=50, textvariable=file_path1)
entry1.grid(row=0,column=1,padx=2,pady=2,sticky='we',columnspan=25)

Label(f2,text="Event Excel File").grid(row=0, column=0, sticky='e') #file2 button
entry2 = Entry(f2, width=50, textvariable=file_path2)
entry2.grid(row=0,column=1,padx=2,pady=2,sticky='we',columnspan=25)



Button(f1, text="Browse", command=browsefunc).grid(row=0, column=27, sticky='ew', padx=8, pady=4)#file1 button
Button(f2, text="Browse", command=browsefunc2).grid(row=0, column=27, sticky='ew', padx=8, pady=4)#file2 button
# Button(f3, text="Browse", command=browsefunc3).grid(row=0, column=27, sticky='ew', padx=8, pady=4)#destination folder button
# Button(f3, text="GetValue", width=32, command=(lambda e=ents: fetch(e))).grid(sticky='ew', padx=10, pady=10)#reconcile button

Button(f4, text="Analyse", width=32, command=lambda: process_file(content)).grid(sticky='ew', padx=10, pady=10)#reconcile button



root.mainloop()