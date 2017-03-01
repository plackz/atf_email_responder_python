## ATF auto responder
# imports
import win32com.client
import pythoncom
import os.path
import datetime
import time

# TODO: assign each message a timestamp "filenumber"
# TODO: remove unwanted characters

# connect to Documentation mailbox.
outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
#inbox = outlook.GetDefaultFolder(6) # works to get items in user inbox
shared_inbox_name = "cvmok"
shared_inbox = outlook.CreateRecipient(shared_inbox_name)
##my_inbox = outlook.GetDefaultFolder(6)
##inbox = my_inbox.Folders.Item("5 weeks")
del_folder = outlook.GetDefaultFolder(3)
inbox = outlook.GetSharedDefaultFolder(shared_inbox, 6)
mailItems = inbox.Items
mailItem = mailItems.GetFirst()
email_counter = 1

while mailItem:

    # create unique tracking ID from datetime
    strFilenum = datetime.datetime.now().strftime("%Y%m%d%H%M%S")

    # delay for 2 seconds
    time.sleep(2) 

    # Check if Subject line has the word "Tracking Number"
    if "Tracking Number" in mailItem.Subject:

        '''
        if there is nothing in the inbox with "Tracking Number" in the subject
        exit and execute the next code block

        if there are no message types then exit and execute the next code block
        '''
    
        strTemp = mailItem.Subject

        # removes unwanted Characters
        strTemp = strTemp.replace(":", "")
        strTemp = strTemp.replace("/", "")
        strTemp = strTemp.replace("\\", "")
        strTemp = strTemp.replace(";", "")
        strTemp = strTemp.replace("?", "")
        strTemp = strTemp.replace("<", "")
        strTemp = strTemp.replace(">", "")
        strTemp = strTemp.replace("|", "")
        strTemp = strTemp.replace("Chr(34)", "")

        # saves message to ATF folder for duplicate request
        #mailItem.Move ("I:\\Quality Control\\After the Fact Documentation\\" & "Additional Request_" & strTemp & ".msg")
        
        #mailItem.SaveAs("C:\\Users\\100355\\" +"Additional Request_" + strTemp + ".msg")   
        # TODO: fix will not delete from mail box.
        #mailItem.Delete
        mailItem.Move(del_folder)

    else:
        strTemp = mailItem.Subject + "_Tracking Number_" + strFilenum
        mailItem.Subject = strTemp
        #itemsUpdated += 1
        mailItem.Save
        '''
        olMailItem = 0x0
        obj = win32com.client.Dispatch("Outlook.Application")
        newMail = obj.CreateItem(olMailItem)

        mailTo = mailItem.Sender
        newMail.Subject = strTemp

        body = '<html><body>' + ' + '</body></html>'

        newMail.HTMLBody = "Thank You for your email. Your request has been received and has been assigned Tracking Number #" & " " & strFilenum & ". Should you have any inquiries on the status of your request, please reference this number."
        newMail.Body = "Thank You for your email. Your request has been received and has been assigned Tracking Number #" & " " & strFilenum & ". Should you have any inquiries on the status of your request, please reference this number."
        newMail.To  = mailTo
        newMail.Send()
        '''
        # need to create escape for instance when there is no mail
        
        '''
        if there is nothing in the inbox with "Tracking Number" in the subject
        exit and execute the next code block

        if there are no message types then exit and execute the next code block
        '''

        # removes unwanted Characters
        strTemp = strTemp.replace(":", "")
        strTemp = strTemp.replace("/", "")
        strTemp = strTemp.replace("\\", "")
        strTemp = strTemp.replace(";", "")
        strTemp = strTemp.replace("?", "")
        strTemp = strTemp.replace("<", "")
        strTemp = strTemp.replace(">", "")
        strTemp = strTemp.replace("|", "")
        strTemp = strTemp.replace("Chr(34)", "")

        # saves file
        #mailItem.SaveAs("I:\\Quality Control\\After the Fact Documentation\\" + strTemp + ".msg")
        #mailItem.Move(del_folder)
        mailItem.SaveAs("C:\\Users\\100355\\" + strTemp + ".msg")
     
    try:
        print(str(mailItem.SenderName))
    except AttributeError:
        print("Error found")

    email_counter += 1
    mailItem = mailItems.GetNext()

# TODO: create a loop to repeat after a certain amount of time.
