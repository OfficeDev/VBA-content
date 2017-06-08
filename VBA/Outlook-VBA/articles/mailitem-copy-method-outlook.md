---
title: MailItem.Copy Method (Outlook)
keywords: vbaol11.chm1321
f1_keywords:
- vbaol11.chm1321
ms.prod: outlook
api_name:
- Outlook.MailItem.Copy
ms.assetid: a9356844-e31e-eb0f-c0f5-a2923ad127db
ms.date: 06/08/2017
---


# MailItem.Copy Method (Outlook)

Creates another instance of an object.


## Syntax

 _expression_ . **Copy**

 _expression_ A variable that represents a **MailItem** object.


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


[MailItem Object](mailitem-object-outlook.md)

