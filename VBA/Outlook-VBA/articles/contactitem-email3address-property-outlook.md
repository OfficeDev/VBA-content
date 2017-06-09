---
title: ContactItem.Email3Address Property (Outlook)
keywords: vbaol11.chm999
f1_keywords:
- vbaol11.chm999
ms.prod: outlook
api_name:
- Outlook.ContactItem.Email3Address
ms.assetid: b0f29077-a06c-a2cf-e873-b9d560d91498
ms.date: 06/08/2017
---


# ContactItem.Email3Address Property (Outlook)

Returns or sets a  **String** representing the e-mail address of the third e-mail entry for the contact. Read/write.


## Syntax

 _expression_ . **Email3Address**

 _expression_ A variable that represents a **ContactItem** object.


## Example

This Visual Basic for Applications (VBA) example sets "someone@example.com" as the e-mail address for the third e-mail entry of a contact.


```vb
Sub CreatePeerContact() 
 
 Dim myItem As Outlook.ContactItem 
 
 
 
 Set myItem = Application.CreateItem(olContactItem) 
 
 myItem.Email3Address = "someone@example.com" 
 
 myItem.Display 
 
End Sub
```


## See also


#### Concepts


[ContactItem Object](contactitem-object-outlook.md)

