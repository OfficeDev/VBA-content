---
title: NoteItem.Copy Method (Outlook)
keywords: vbaol11.chm1493
f1_keywords:
- vbaol11.chm1493
ms.prod: outlook
api_name:
- Outlook.NoteItem.Copy
ms.assetid: 5d89217e-2595-64e2-a619-afb5a7120f8a
ms.date: 06/08/2017
---


# NoteItem.Copy Method (Outlook)

Creates another instance of an object.


## Syntax

 _expression_ . **Copy**

 _expression_ An expression that returns a **NoteItem** object.


### Return Value

A  **[NoteItem](noteitem-object-outlook.md)** object that represents a copy of the specified note.


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


[NoteItem Object](noteitem-object-outlook.md)

