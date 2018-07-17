---
title: MailItem.Importance Property (Outlook)
keywords: vbaol11.chm1306
f1_keywords:
- vbaol11.chm1306
ms.prod: outlook
api_name:
- Outlook.MailItem.Importance
ms.assetid: 77de74c9-e910-e021-1015-6e65f3ead3df
ms.date: 06/08/2017
---


# MailItem.Importance Property (Outlook)

Returns or sets an  **[OlImportance](olimportance-enumeration-outlook.md)** constant indicating the relative importance level for the Outlook item. Read/write.


## Syntax

 _expression_ . **Importance**

 _expression_ A variable that represents a **MailItem** object.


## Remarks

This property corresponds to the MAPI property  **PidTagImportance** .


## Example

This Visual Basic for Applications (VBA) example checks if the item displayed in the topmost inspector is sent by 'Dan Wilson' with 'High' importance. If it is, then it displays a message box to the user. Before running this example, replace 'Dan Wilson' with a valid name in your address book.


```vb
Sub CheckSenderName 
 
 Dim myItem As Outlook.MailItem 
 
 
 
 Set myItem = Application.ActiveInspector.CurrentItem 
 
 If myItem.Importance = 2 And _ 
 
 myItem.SenderName = "Dan Wilson" Then 
 
 MsgBox "This message is sent by your manager with High importance." 
 
 End If 
 
End Sub
```


## See also


#### Concepts


[MailItem Object](mailitem-object-outlook.md)

