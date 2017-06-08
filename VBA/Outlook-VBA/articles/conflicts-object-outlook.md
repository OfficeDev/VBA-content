---
title: Conflicts Object (Outlook)
keywords: vbaol11.chm399
f1_keywords:
- vbaol11.chm399
ms.prod: outlook
api_name:
- Outlook.Conflicts
ms.assetid: c4e1c060-519a-a6d1-8fb2-c7dfa1e3e66f
ms.date: 06/08/2017
---


# Conflicts Object (Outlook)

Contains a collection of  **[Conflict](conflict-object-outlook.md)** objects that represent all Microsoft Outlook items that are in conflict with a particular Outlook item.


## Remarks

Use the  **[Conflicts](mailitem-conflicts-property-outlook.md)** property of any Outlook item, such as **[MailItem](mailitem-object-outlook.md)**, to return the **Conflicts** object.

Use the  **[Count](conflicts-count-property-outlook.md)** property of the **Conflicts** object to determine if the item is invloved in a conflict. A non-zero value indicates conflict.

Use the  **[Item](conflicts-item-method-outlook.md)** method to retrieve a particular conflict item from the **Conflicts** collection object.

Use the  **[GetFirst](conflicts-getfirst-method-outlook.md)**, **[GetNext](conflicts-getnext-method-outlook.md)**, **[GetPrevious](conflicts-getprevious-method-outlook.md)**, and **[GetLast](conflicts-getlast-method-outlook.md)** methods to traverse the **Conflicts** collection.


## Example

The following Microsoft Visual Basic for Applications (VBA) example uses the  **Count** property of the **Conflicts** object to determine if the item is involved in any conflict. To run this example, make sure an e-mail item is open in the active window.


```
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


## Methods



|**Name**|
|:-----|
|[GetFirst](conflicts-getfirst-method-outlook.md)|
|[GetLast](conflicts-getlast-method-outlook.md)|
|[GetNext](conflicts-getnext-method-outlook.md)|
|[GetPrevious](conflicts-getprevious-method-outlook.md)|
|[Item](conflicts-item-method-outlook.md)|

## Properties



|**Name**|
|:-----|
|[Application](conflicts-application-property-outlook.md)|
|[Class](conflicts-class-property-outlook.md)|
|[Count](conflicts-count-property-outlook.md)|
|[Parent](conflicts-parent-property-outlook.md)|
|[Session](conflicts-session-property-outlook.md)|

## See also


#### Other resources


[Outlook Object Model Reference](http://msdn.microsoft.com/library/73221b13-d8d8-99b8-3394-b95dbbfd5ddc%28Office.15%29.aspx)
