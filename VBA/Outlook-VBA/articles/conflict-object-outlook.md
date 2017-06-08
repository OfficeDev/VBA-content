---
title: Conflict Object (Outlook)
keywords: vbaol11.chm410
f1_keywords:
- vbaol11.chm410
ms.prod: outlook
api_name:
- Outlook.Conflict
ms.assetid: a7c8f12a-08ba-9fff-60b8-a02d1c7f6f33
ms.date: 06/08/2017
---


# Conflict Object (Outlook)

Represents an Outlook item that is in conflict with another Outlook item.


## Remarks

 Each Outlook item has a **[Conflicts](conflicts-object-outlook.md)** collection object associated with it that represents all the items that are in conflict with that item.

Use the  **[Item](conflicts-item-method-outlook.md)** method to retrieve a particular **Conflict** object from the **Conflicts** collection object, for example:


## Example

The following Visual Basic for Applications (VBA) example retrieves a  **Conflict** object from the **Conflicts** collection object.


```
Set myConflictItem = myConflicts.Item(1)
```


## Properties



|**Name**|
|:-----|
|[Application](conflict-application-property-outlook.md)|
|[Class](conflict-class-property-outlook.md)|
|[Item](conflict-item-property-outlook.md)|
|[Name](conflict-name-property-outlook.md)|
|[Parent](conflict-parent-property-outlook.md)|
|[Session](conflict-session-property-outlook.md)|
|[Type](conflict-type-property-outlook.md)|

## See also


#### Other resources


[Outlook Object Model Reference](http://msdn.microsoft.com/library/73221b13-d8d8-99b8-3394-b95dbbfd5ddc%28Office.15%29.aspx)
