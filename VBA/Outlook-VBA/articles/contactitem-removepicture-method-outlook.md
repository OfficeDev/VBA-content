---
title: ContactItem.RemovePicture Method (Outlook)
keywords: vbaol11.chm1091
f1_keywords:
- vbaol11.chm1091
ms.prod: outlook
api_name:
- Outlook.ContactItem.RemovePicture
ms.assetid: a67d9d39-1697-0780-b52f-a3cc463f60d9
ms.date: 06/08/2017
---


# ContactItem.RemovePicture Method (Outlook)

Removes a picture from a  **Contacts** item.


## Syntax

 _expression_ . **RemovePicture**

 _expression_ A variable that represents a **ContactItem** object.


## Example

The following Microsoft Visual Basic for Applications (VBA) example prompts the user to specify a name of a contact and removes the picture from the contact item. If a picture does not exist for the contact, the example displays a message to the user.


```vb
Sub RemovePictureFromContact() 
 
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
 
 If myContactItem.HasPicture = False Then 
 
 MsgBox "The contact does not have a picture associated with it." 
 
 Else 
 
 myContactItem.RemovePicture 
 
 myContactItem.Save 
 
 myContactItem.Display 
 
 End If 
 
End Sub
```


## See also


#### Concepts


[ContactItem Object](contactitem-object-outlook.md)

