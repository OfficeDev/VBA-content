---
title: ContactItem.Email1Address Property (Outlook)
keywords: vbaol11.chm991
f1_keywords:
- vbaol11.chm991
ms.prod: outlook
api_name:
- Outlook.ContactItem.Email1Address
ms.assetid: 0bd407bc-21a9-16e6-709d-383cb79b4d6e
ms.date: 06/08/2017
---


# ContactItem.Email1Address Property (Outlook)

Returns or sets a  **String** representing the e-mail address of the first e-mail entry for the contact. Read/write.


## Syntax

 _expression_ . **Email1Address**

 _expression_ A variable that represents a **ContactItem** object.


## Example

This Visual Basic for Applications (VBA) example sets "someone@example.com" as the e-mail address for the first e-mail entry of a contact.


```vb
Sub CreatePeerContact() 
 
 Dim myItem As Outlook.ContactItem 
 
 
 
 Set myItem = Application.CreateItem(olContactItem) 
 
 myItem.Email1Address = "someone@example.com" 
 
 myItem.Display 
 
End Sub
```


## See also


#### Concepts


[ContactItem Object](contactitem-object-outlook.md)

