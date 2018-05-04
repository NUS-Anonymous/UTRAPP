
from tkinter import *
import tkinter.filedialog

#import tkinter, filedialog

fields = 'Main Excel File', 'Event Excel File', 'From Event', 'Email Login', 'Email Password','Email Content'

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
      ent.pack(side=RIGHT, expand=YES, fill=X)
      entries.append((field, ent))
   return entries

def loadtemplate(self): 
        filename = tkinter.filedialog.askopenfilename(filetypes = (("Template files", "*.tplate")
                                                             ,("HTML files", "*.html;*.htm")
                                                             ,("All files", "*.*") ))
        if filename: 
            try: 
                self.settings["template"].set(filename)
            except: 
                tkMessageBox.showerror("Open Source File", "Failed to read file \n'%s'"%filename)
                return

def load_file():
        fname = tkinter.filedialog.askopenfilename(filetypes=(("Excel", "*.xlsx"),
                                           ("All files", "*.*") ))
        if fname:
            try:
                print("""here it comes: self.settings["template"].set(fname)""")
            except:                     # <- naked except is a bad idea
                showerror("Open Source File", "Failed to read file\n'%s'" % fname)
            return

def browsefunc():
      filename = tkinter.filedialog.askopenfilename()
      pathlabel.config(text=filename)
      # write(filename)



if __name__ == '__main__':
   root = Tk()
   browsebutton = Button(root, text="Browse", command=browsefunc)
   browsebutton.pack(side=LEFT, padx=5, pady=5)

   pathlabel = Label(root)
   pathlabel.pack(fill=X)
   
   ents = makeform(root, fields)
   root.bind('<Return>', (lambda event, e=ents: fetch(e)))   
   b1 = Button(root, text='Analyse',
          command=(lambda e=ents: fetch(e)))
   b1.pack(side=LEFT, padx=5, pady=5)
   b2 = Button(root, text='Send Email',
          command=(lambda e=ents: fetch(e)))
   b2.pack(side=LEFT, padx=5, pady=5)
   b3 = Button(root, text='Quit', command=root.quit)
   b3.pack(side=LEFT, padx=5, pady=5)
   # b4 = Button(root, text="Browse", command=load_file(), width=10)
   # b4.pack(side=LEFT, padx=5, pady=5)
   w = Label(root, text="Red", bg="red", fg="white")
   w.pack()
  


   root.mainloop()