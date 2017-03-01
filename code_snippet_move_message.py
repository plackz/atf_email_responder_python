import win32com.client
import time

outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
InboxName = outlook.CreateRecipient("cvmok")
Inbox = outlook.GetSharedDefaultFolder(InboxName, 6)
for aItem in Inbox.Items:
    fileNum = time.strftime("%Y%m%d%H%M%S")
    
    tempSubj = aItem.Subject
    
    time.sleep(2)
    
    if "Tracking Number" in str(aItem.Subject):
        #print("Subject: "+ "Additional Request_" + tempSubj)
        newTempSubj = ("Subject: " + "Additional Request_" + tempSubj)
        print(newTempSubj)
        aItem.SaveAs(r"C:\\TempData\\" + newTempSubj + ".msg") # olMSG= "3" for outlook
    else:
        print("Subject: "+ tempSubj + "_Tracking Number_" + fileNum)
