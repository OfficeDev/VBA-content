---
title: AppointmentItem.Copy Method (Outlook)
keywords: vbaol11.chm869
f1_keywords:
- vbaol11.chm869
ms.prod: outlook
api_name:
- Outlook.AppointmentItem.Copy
ms.assetid: 947f1cfd-f60c-a47e-ba4d-3ffde8c13c91
ms.date: 06/08/2017
---


# AppointmentItem.Copy Method (Outlook)

Creates another instance of an object.


## Syntax

 _expression_ . **Copy**

 _expression_ A variable that represents an **AppointmentItem** object.


## Example

This Visual Basic for Applications example creates an e-mail message, sets the  **Subject** to "Speeches", uses the **Copy** method to copy it, then moves the copy into a newly created e-mail folder named "Saved Mail" within the Inbox folder.


```vb
Sub CopyItem() 
 
 Dim myNameSpace As Outlook.NameSpace 
 
 Dim myFolder As Outlook.Folder 
 
 Dim myNewFolder As Outlook.Folder 
 
 Dim myItem As Outlook.MailItem 
 
 Dim myCopiedItem As Outlook.MailItem 
 
 
 
 Set myNameSpace = Application.GetNamespace("MAPI") 
 
 Set myFolder = myNameSpace.GetDefaultFolder(olFolderInbox) 
 
 Set myNewFolder = myFolder.Folders.Add("Saved Mail", olFolderDrafts) 
 
 Set myItem = Application.CreateItem(olMailItem) 
 
 myItem.Subject = "Speeches" 
 
 Set myCopiedItem = myItem.Copy 
 
 myCopiedItem.Move myNewFolder 
 
End Sub
```


## See also


#### Concepts


[AppointmentItem Object](appointmentitem-object-outlook.md)

