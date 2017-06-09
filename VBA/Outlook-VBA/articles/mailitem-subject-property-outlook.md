---
title: MailItem.Subject Property (Outlook)
keywords: vbaol11.chm1317
f1_keywords:
- vbaol11.chm1317
ms.prod: outlook
api_name:
- Outlook.MailItem.Subject
ms.assetid: 5f3e465d-ac2b-a573-0e85-1134e65df017
ms.date: 06/08/2017
---


# MailItem.Subject Property (Outlook)

Returns or sets a  **String** indicating the subject for the Outlook item. Read/write.


## Syntax

 _expression_ . **Subject**

 _expression_ A variable that represents a **[MailItem](mailitem-object-outlook.md)** object.


## Remarks

This property corresponds to the MAPI property  **PidTagSubject** . The **Subject** property is the default property for Outlook items.


## Example

This Microsoft Visual Basic for Applications (VBA) example creates a new e-mail message, uses the  **[Add](recipients-add-method-outlook.md)** method to add "Dan Wilson" as a **[To](mailitem-to-property-outlook.md)** recipient, sets the **Subject** property, and displays the message.


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


[Send an E-mail Given the SMTP Address of an Account](send-an-e-mail-given-the-smtp-address-of-an-account-outlook.md)


