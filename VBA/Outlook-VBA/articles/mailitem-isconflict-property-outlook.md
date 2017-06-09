---
title: MailItem.IsConflict Property (Outlook)
keywords: vbaol11.chm1377
f1_keywords:
- vbaol11.chm1377
ms.prod: outlook
api_name:
- Outlook.MailItem.IsConflict
ms.assetid: 648e6b53-81fb-03ec-0029-edbdd05c663b
ms.date: 06/08/2017
---


# MailItem.IsConflict Property (Outlook)

Returns a  **Boolean** that determines if the item is in conflict. Read-only.


## Syntax

 _expression_ . **IsConflict**

 _expression_ A variable that represents a **MailItem** object.


## Remarks

Whether or not an item is in conflict is determined by the state of the application. For example, when a user is offline and tries to access an online folder the action will fail. In this scenario, the  **IsConflict** property will return **True** .

If  **True** , the specified item is in conflict.


## Example

The following Microsoft Visual Basic for Applications (VBA) example creates a new mail item and attempts to send it. If the  **IsConflict** property returns **True** , the item will not be sent.


```vb
Sub NewMail() 
 
 'Creates and tries to send a new e-mail message. 
 
 Dim objNewMail As Outlook.MailItem 
 
 
 
 Set objNewMail = Application.CreateItem(olMailItem) 
 
 objNewMail.Body = _ 
 
 "This e-mail message was created automatically on " &; Now 
 
 objNewMail.To = "Jeff Smith" 
 
 If objNewMail.IsConflict = False Then 
 
 objNewMail.Send 
 
 Else 
 
 MsgBox "Conflict: Cannot send mail item." 
 
 End If 
 
 Set olApp = Nothing 
 
 Set objNewMail = Nothing 
 
End Sub
```


## See also


#### Concepts


[MailItem Object](mailitem-object-outlook.md)

