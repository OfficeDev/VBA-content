---
title: ContactItem.Email2Address Property (Outlook)
keywords: vbaol11.chm995
f1_keywords:
- vbaol11.chm995
ms.prod: outlook
api_name:
- Outlook.ContactItem.Email2Address
ms.assetid: 1656eb41-55b3-50f7-7351-b287e07bcac0
ms.date: 06/08/2017
---


# ContactItem.Email2Address Property (Outlook)

Returns or sets a  **String** representing the e-mail address of the second e-mail entry for the contact. Read/write.


## Syntax

 _expression_ . **Email2Address**

 _expression_ A variable that represents a **ContactItem** object.


## Example

This Visual Basic for Applications (VBA) example sets "someone@example.com" as the e-mail address for the second e-mail entry of a contact.


```vb
Sub CreatePeerContact() 
 
 Dim myItem As Outlook.ContactItem 
 
 
 
 Set myItem = Application.CreateItem(olContactItem) 
 
 myItem.Email2Address = "someone@example.com" 
 
 myItem.Display 
 
End Sub
```


## See also


#### Concepts


[ContactItem Object](contactitem-object-outlook.md)

