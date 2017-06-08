---
title: MailItem.Conflicts Property (Outlook)
keywords: vbaol11.chm1382
f1_keywords:
- vbaol11.chm1382
ms.prod: outlook
api_name:
- Outlook.MailItem.Conflicts
ms.assetid: 2c93c2a2-4f2f-17af-cba3-91620b3d9c0f
ms.date: 06/08/2017
---


# MailItem.Conflicts Property (Outlook)

Return the  **[Conflicts](conflicts-object-outlook.md)** object that represents the items that are in conflict for any Outlook item object. Read-only.


## Syntax

 _expression_ . **Conflicts**

 _expression_ A variable that represents a **MailItem** object.


## Example

The following Microsoft Visual Basic for Applications (VBA) example uses the  **[Count](conflicts-count-property-outlook.md)** property of the **Conflicts** object to determine if the item is involved in any conflict. To run this example, make sure a mail item is open in the active window.


```vb
Sub CheckConflicts() 
 
 Dim myItem As Outlook.MailItem 
 
 Dim myConflicts As Outlook.Conflicts 
 
 
 
 Set myItem = Application.ActiveInspector.CurrentItem 
 
 Set myConflicts = myItem.Conflicts 
 
 If (myConflicts.Count > 0) Then 
 
 MsgBox ("This item is involved in a conflict.") 
 
 Else 
 
 MsgBox ("This item is not involved in any conflicts.") 
 
 End If 
 
End Sub
```


## See also


#### Concepts


[MailItem Object](mailitem-object-outlook.md)

