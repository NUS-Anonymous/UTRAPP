import pandas as pd
import numpy as np

Database= "/home/ubun/Desktop/UTRApp/UTR.xlsx"
Event = "/home/ubun/Desktop/UTRApp/ML.xlsx"

eventName = "Event1"
mergeOn='Name'
totalFrom='E1'





xl_fileDB = pd.ExcelFile(Database)
DB = {sheet_name: xl_fileDB.parse(sheet_name) 
          for sheet_name in xl_fileDB.sheet_names}

#Get the first Sheet Name
DBsheet = xl_fileDB.sheet_names[0].encode('ascii','ignore')

xl_fileE = pd.ExcelFile(Event)
Event = {sheet_name: xl_fileE.parse(sheet_name) 
          for sheet_name in xl_fileE.sheet_names}

Eventsheet = xl_fileE.sheet_names[0].encode('ascii','ignore')

#sum from column index
df = xl_fileDB.parse(xl_fileDB.sheet_names[0])
index=0
if totalFrom!="":
	index=df.columns.get_loc(totalFrom)

Event[Eventsheet][eventName]=1
s1 = pd.merge(DB[DBsheet], Event[Eventsheet], how='left', on=[mergeOn])
s1['Total']=0
s1['Total']=s1.iloc[:,index:].sum(axis=1)

writer = pd.ExcelWriter('./output.xlsx')
s1.to_excel(writer,'Sheet1')
writer.save()


# #Sending from GMAil
# SERVER = "smtp.google.com"
# FROM = "johnDoe@gmail.com"
# TO = ["JaneDoe@gmail.com"] # must be a list

# SUBJECT = "Hello!"
# TEXT = "This is a test of emailing through smtp in google."

# # Prepare actual message
# message = """From: %s\r\nTo: %s\r\nSubject: %s\r\n\

# %s
# """ % (FROM, ", ".join(TO), SUBJECT, TEXT)

# # Send the mail
# import smtplib
# server = smtplib.SMTP(SERVER)
# server.login("MrDoe", "PASSWORD")
# server.sendmail(FROM, TO, message)
# server.quit()
# #Sending from Outlook
# SERVER = "your.mail.server"
# FROM = "yourEmail@yourAddress.com"
# TO = ["listOfEmails"] # must be a list

# SUBJECT = "Subject"
# TEXT = "Your Text"

# # Prepare actual message
# message = """From: %s\r\nTo: %s\r\nSubject: %s\r\n\

# %s
# """ % (FROM, ", ".join(TO), SUBJECT, TEXT)

# # Send the mail
# import smtplib
# server = smtplib.SMTP(SERVER)
# server.sendmail(FROM, TO, message)
# server.quit()
# #