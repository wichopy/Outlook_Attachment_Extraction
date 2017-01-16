# -*- coding: utf-8 -*-
"""
Created on Wed May 11 15:37:40 2016
## use this for pulling email attachments from Outlook. 
Play around with the GetDefaultFolder numbers to get to the folder you want to in Outlook. 
From there, run the for loop to get all the emails.
Modify the savefileas method to choose your saved folder
@author: chouw
"""

import win32com.client
import os
#This chunk of code is for navigating to your desired folder where all the attachments you want are.
#Display all mailboxes
outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
#select mailbox
inbox = outlook.GetDefaultFolder(6)
#select folder you're interested in in mailbox.
allfolds = outlook.Folders
PSfold = allfolds[2]
print PSfold.folders[18].name #verify this is the folder I want

#This chunk if for extracting the attachments.
# isolate for emails.
emails = PSfold.folders[18].items
attachments = []
count = 0
for i in xrange(emails.count):
	#extract attachments.
        attachments.append(emails[i].attachments)
        count = count +1
       
        print ("saving to mydocuments:")
        print attachments[i].item(1).filename
        #desired saved folder.
        attachments[i].item(1).saveasfile("C:/Users/chouw/Documents/OLAttachments/")
