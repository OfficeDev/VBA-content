---
title: MailItem.SenderName Property (Outlook)
keywords: vbaol11.chm1357
f1_keywords:
- vbaol11.chm1357
ms.prod: outlook
api_name:
- Outlook.MailItem.SenderName
ms.assetid: e3c133e6-c7a8-9004-969d-aa2a466f8486
ms.date: 06/08/2017
---


# MailItem.SenderName Property (Outlook)

Returns a  **String** indicating the display name of the sender for the Outlook item. Read-only.


## Syntax

 _expression_ . **SenderName**

 _expression_ A variable that represents a **MailItem** object.


## Remarks

This property corresponds to the MAPI property  **PidTagSenderName** .

If you wish to retrieve the fully qualified e-mail address of the sender, use the  **[SenderEmailAddress](mailitem-senderemailaddress-property-outlook.md)** property.


## Example

This Visual Basic for Applications (VBA) example checks if the item displayed in the topmost inspector is sent by 'Dan Wilson' with 'High' importance. If it is, then it displays a message box to the user. Before running this example, replace 'Dan Wilson' with a valid name in your address book.


```vb
Sub CheckSenderName 
 
 Dim myItem As Outlook.MailItem 
 
 
 
 Set myItem = Application.ActiveInspector.CurrentItem 
 
 If myItem.Importance = 2 And myItem.SenderName = _ 
 
 "Dan Wilson" Then 
 
 MsgBox "This message is sent by your manager with High importance." 
 
 End If 
 
End Sub
```


## See also


#### Concepts


[MailItem Object](mailitem-object-outlook.md)

