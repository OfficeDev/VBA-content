---
title: OutlookBarGroups Object (Outlook)
keywords: vbaol11.chm3002
f1_keywords:
- vbaol11.chm3002
ms.prod: outlook
api_name:
- Outlook.OutlookBarGroups
ms.assetid: bb5fef46-b15a-51c3-0adf-f94e9da6c921
ms.date: 06/08/2017
---


# OutlookBarGroups Object (Outlook)

Contains a set of  **[OutlookBarGroup](outlookbargroup-object-outlook.md)** objects representing all groups in the Outlook Bar.


## Remarks

Use the  **[Groups](outlookbarstorage-groups-property-outlook.md)** property to return the **OutlookBarGroups** object from the **[OutlookBarStorage](outlookbarstorage-object-outlook.md)** object.

Use  **Groups** ( _index_ ), where _index_ is the name of an available group, to return a single **OutlookBarGroup** object.


## Example

The following Visual Basic for Applications (VBA) example retrieves the  **OutlookBarGroups** collection from an **OutlookBarStorage** object.


```
Set myGroups = myOutlookBarStorage.Groups
```


## Events



|**Name**|
|:-----|
|[BeforeGroupAdd](outlookbargroups-beforegroupadd-event-outlook.md)|
|[BeforeGroupRemove](outlookbargroups-beforegroupremove-event-outlook.md)|
|[GroupAdd](outlookbargroups-groupadd-event-outlook.md)|

## Methods



|**Name**|
|:-----|
|[Add](outlookbargroups-add-method-outlook.md)|
|[Item](outlookbargroups-item-method-outlook.md)|
|[Remove](outlookbargroups-remove-method-outlook.md)|

## Properties



|**Name**|
|:-----|
|[Application](outlookbargroups-application-property-outlook.md)|
|[Class](outlookbargroups-class-property-outlook.md)|
|[Count](outlookbargroups-count-property-outlook.md)|
|[Parent](outlookbargroups-parent-property-outlook.md)|
|[Session](outlookbargroups-session-property-outlook.md)|

## See also


#### Other resources


[Outlook Object Model Reference](http://msdn.microsoft.com/library/73221b13-d8d8-99b8-3394-b95dbbfd5ddc%28Office.15%29.aspx)
