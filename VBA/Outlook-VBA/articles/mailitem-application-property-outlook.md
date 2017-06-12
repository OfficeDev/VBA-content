---
title: MailItem.Application Property (Outlook)
keywords: vbaol11.chm1290
f1_keywords:
- vbaol11.chm1290
ms.prod: outlook
api_name:
- Outlook.MailItem.Application
ms.assetid: d71cb356-f3ae-ab08-4209-1dac0c2b8fdf
ms.date: 06/08/2017
---


# MailItem.Application Property (Outlook)

Returns an  **[Application](application-object-outlook.md)** object that represents the parent Outlook application for the object. Read-only.


## Syntax

 _expression_ . **Application**

 _expression_ A variable that represents a **MailItem** object.


## Example

This Visual Basic for Applications (VBA) example uses the  **Application** property to access Outlook, creates a new **[MailItem](mailitem-object-outlook.md)** and displays the version of Outlook used to create the item.


```vb
Sub CreateMailItem() 
 
 Dim myItem As Outlook.MailItem 
 
 
 
 Set myItem = Application.CreateItem(olMailItem) 
 
 MsgBox myItem.Application.Version 
 
End Sub
```


## See also


#### Concepts


[MailItem Object](mailitem-object-outlook.md)

