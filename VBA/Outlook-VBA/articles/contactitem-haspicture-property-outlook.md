---
title: ContactItem.HasPicture Property (Outlook)
keywords: vbaol11.chm1092
f1_keywords:
- vbaol11.chm1092
ms.prod: outlook
api_name:
- Outlook.ContactItem.HasPicture
ms.assetid: 5e835af9-bcee-692d-f486-5f8a4a0efa1a
ms.date: 06/08/2017
---


# ContactItem.HasPicture Property (Outlook)

Returns a  **Boolean** value that is **True** if a **Contacts** item has a picture associated with it. Read-only


## Syntax

 _expression_ . **HasPicture**

 _expression_ A variable that represents a **ContactItem** object.


## Example

The following Microsoft Visual Basic for Applications (VBA) example prompts the user to specify the name of a contact and the file name containing a picture of the contact, and then adds the picture to the contact item. If a picture already exists for the contact item, the example prompts the user to specify if the existing picture should be overwritten by the new file.


```vb
Sub AddPictureToAContact() 
 
 Dim myNms As Outlook.NameSpace 
 
 Dim myFolder As Outlook.Folder 
 
 Dim myContactItem As Outlook.ContactItem 
 
 Dim strName As String 
 
 Dim strPath As String 
 
 Dim strPrompt As String 
 
 
 
 Set myNms = Application.GetNamespace("MAPI") 
 
 Set myFolder = myNms.GetDefaultFolder(olFolderContacts) 
 
 strName = InputBox("Type the name of the contact: ") 
 
 Set myContactItem = myFolder.Items(strName) 
 
 If myContactItem.HasPicture = True Then 
 
 strPrompt = MsgBox("The contact already has a picture associated with it. Do you want to overwrite the existing picture?", vbYesNo) 
 
 If strPrompt = vbNo Then 
 
 Exit Sub 
 
 End If 
 
 End If 
 
 strPath = InputBox("Type the file name for the contact: ") 
 
 myContactItem.AddPicture (strPath) 
 
 myContactItem.Save 
 
 myContactItem.Display 
 
End Sub
```


## See also


#### Concepts


[ContactItem Object](contactitem-object-outlook.md)

