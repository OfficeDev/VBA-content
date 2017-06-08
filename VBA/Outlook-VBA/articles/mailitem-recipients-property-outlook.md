---
title: MailItem.Recipients Property (Outlook)
keywords: vbaol11.chm1347
f1_keywords:
- vbaol11.chm1347
ms.prod: outlook
api_name:
- Outlook.MailItem.Recipients
ms.assetid: 58897f66-8a6a-e1a9-7e3b-5a84624f899d
ms.date: 06/08/2017
---


# MailItem.Recipients Property (Outlook)

Returns a  **[Recipients](recipients-object-outlook.md)** collection that represents all the recipients for the Outlook item. Read-only.


## Syntax

 _expression_ . **Recipients**

 _expression_ A variable that represents a **[MailItem](mailitem-object-outlook.md)** object.


## Remarks

A recipient can be specified by a string representing the recipient's display name, alias, or full SMTP e-mail address.


## Example

This Visual Basic for Applications (VBA) example creates a new e-mail message, uses the  **[Add](recipients-add-method-outlook.md)** method to add "Dan Wilson" as a **[To](mailitem-to-property-outlook.md)** recipient, and displays the message.


```vb
Sub CreateStatusReportToBoss() 
 
 Dim myItem As Outlook.MailItem 
 
 Dim myRecipient As Outlook.Recipient 
 
 
 
 Set myItem = Application.CreateItem(olMailItem) 
 
 Set myRecipient = myItem.Recipients.Add("Dan Wilson") 
 
 myItem.Subject = "Status Report" 
 
 myItem.Display 
 
End Sub
```


## See also


#### Concepts


[MailItem Object](mailitem-object-outlook.md)
#### Other resources


[How to: Send an E-mail Given the SMTP Address of an Account](http://msdn.microsoft.com/library/97406049-f63a-0c1d-9b3f-57bf48afc4be%28Office.15%29.aspx)

[How to: Send an E-mail Given the SMTP Address of an Account](http://msdn.microsoft.com/library/5e5f707d-8771-bd5f-945b-58537732d99a%28Office.15%29.aspx)

