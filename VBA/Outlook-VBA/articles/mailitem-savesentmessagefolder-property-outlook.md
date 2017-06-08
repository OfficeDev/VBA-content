---
title: MailItem.SaveSentMessageFolder Property (Outlook)
keywords: vbaol11.chm1356
f1_keywords:
- vbaol11.chm1356
ms.prod: outlook
api_name:
- Outlook.MailItem.SaveSentMessageFolder
ms.assetid: ab36ae3b-6c6d-842b-dbb4-88c37d8e7874
ms.date: 06/08/2017
---


# MailItem.SaveSentMessageFolder Property (Outlook)

Returns or sets a  **[Folder](folder-object-outlook.md)** object that represents the folder in which a copy of the e-mail message will be saved after being sent. Read/write.


## Syntax

 _expression_ . **SaveSentMessageFolder**

 _expression_ A variable that represents a **MailItem** object.


## Example

This Visual Basic for Applications (VBA) example sends a reply to Dan Wilson and sets the  `SaveMyPersonalItems` folder as the folder in which a copy of the item will be saved after being sent. To run this example without errors, make sure a mail item is open in the active inspector window and replace 'Dan Wilson' with a valid recipient name.


```vb
Sub SetSentFolder() 
 
 Dim myItem As Outlook.MailITem 
 
 Dim myResponse As Outlook.MailITem 
 
 Dim mpfInbox As Outlook.Folder 
 
 Dim mpf As Outlook.Folder 
 
 
 
 Set mpfInbox = Application.Session.GetDefaultFolder(olFolderInbox) 
 
 Set mpf = mpfInbox.Folders.Add("SaveMyPersonalItems") 
 
 Set myItem = Application.ActiveInspector.CurrentItem 
 
 Set myResponse = myItem.Reply 
 
 myResponse.Display 
 
 myResponse.To = "Dan Wilson" 
 
 Set myResponse.SaveSentMessageFolder = mpf 
 
 myResponse.Send 
 
End Sub
```


## See also


#### Concepts


[MailItem Object](mailitem-object-outlook.md)

