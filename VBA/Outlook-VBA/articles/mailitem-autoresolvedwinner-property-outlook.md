---
title: MailItem.AutoResolvedWinner Property (Outlook)
keywords: vbaol11.chm1381
f1_keywords:
- vbaol11.chm1381
ms.prod: outlook
api_name:
- Outlook.MailItem.AutoResolvedWinner
ms.assetid: 3c0ccbd5-47a6-7a0c-a488-037c48fc1958
ms.date: 06/08/2017
---


# MailItem.AutoResolvedWinner Property (Outlook)

Returns a  **Boolean** that determines if the item is a winner of an automatic conflict resolution. Read-only.


## Syntax

 _expression_ . **AutoResolvedWinner**

 _expression_ A variable that represents a **MailItem** object.


## Remarks

A value of  **False** does not necessarily indicate that the item is a loser of an automatic conflict resolution. The item could be in conflict with another item.

If an item has  **[Conflicts.Count](conflicts-count-property-outlook.md)** of its **[MailItem.Conflicts](mailitem-conflicts-property-outlook.md)** property greater than zero and if its **AutoResolvedWinner** property is **True** , it is a winner of an automatic conflict resolution. On the other hand, if the item is in conflict and has its **AutoResolvedWinner** property as **False** , it is a loser in an automatic conflict resolution.


## Example

The following Microsoft Visual Basic for Applications (VBA) example used the  **AutoResolvedWinner** property to determine if an item is a winner or loser in an automatic conflict resolution. To run this example, make sure an e-mail item is open in the active window.


```vb
Sub ConflictStatus() 
 
 Dim mail As Outlook.MailItem 
 
 Set mail = Application.ActiveInspector.CurrentItem 
 
 If mail.Conflicts.Count > 0 Then 
 
 If mail.AutoResolvedWinner = True Then 
 
 MsgBox "This item is a winner in an automatic conflict resolution." 
 
 Else 
 
 MsgBox "This item is a loser in an automatic conflict resolution." 
 
 End If 
 
 Else 
 
 MsgBox "This item is not in conflict with any item." 
 
 End If 
 
End Sub
```


## See also


#### Concepts


[MailItem Object](mailitem-object-outlook.md)

