---
title: ContactItem.Email2AddressType Property (Outlook)
keywords: vbaol11.chm996
f1_keywords:
- vbaol11.chm996
ms.prod: outlook
api_name:
- Outlook.ContactItem.Email2AddressType
ms.assetid: 09e1448e-87d7-5040-a13f-ae8d7ae67cb9
ms.date: 06/08/2017
---


# ContactItem.Email2AddressType Property (Outlook)

Returns or sets a  **String** representing the address type (such as EX or SMTP) of the second e-mail entry for the contact. Read/write.


## Syntax

 _expression_ . **Email2AddressType**

 _expression_ A variable that represents a **ContactItem** object.


## Remarks

This is a free-form text field, but it must match the actual type of an existing e-mail transport.


## Example

This Visual Basic for Applications (VBA) example sets "SMTP" as the address type for the second e-mail entry of a contact.


```vb
Sub SetType() 
 
 Dim myItem As Outlook.ContactItem 
 
 
 
 Set myItem = Application.CreateItem(olContactItem) 
 
 myItem.Email2Address = "someone@example.com" 
 
 myItem.Email2AddressType = "SMTP" 
 
 myItem.Display 
 
End Sub
```


## See also


#### Concepts


[ContactItem Object](contactitem-object-outlook.md)

