---
title: OutlookBarStorage Object (Outlook)
keywords: vbaol11.chm367
f1_keywords:
- vbaol11.chm367
ms.prod: outlook
api_name:
- Outlook.OutlookBarStorage
ms.assetid: e6dc8dc0-bae4-f59b-c991-1421b280de38
ms.date: 06/08/2017
---


# OutlookBarStorage Object (Outlook)

Represents the storage for objects in the Outlook Bar.


## Remarks

Use the  **[Contents](outlookbarpane-contents-property-outlook.md)** property of an **[OutlookBarPane](outlookbarpane-object-outlook.md)** object to retrieve the **OutlookBarStorage** object for the Outlook Bar.

Use the  **[Groups](outlookbarstorage-groups-property-outlook.md)** property to retrieve the **[OutlookBarGroups](outlookbargroups-object-outlook.md)** object for the Outlook Bar.


## Example

The following example retrieves an  **OutlookBarStorage** object by name.


```
Set myOLBarStorage = myPanes.Item("OutlookBar").Contents
```


## Properties



|**Name**|
|:-----|
|[Application](outlookbarstorage-application-property-outlook.md)|
|[Class](outlookbarstorage-class-property-outlook.md)|
|[Groups](outlookbarstorage-groups-property-outlook.md)|
|[Parent](outlookbarstorage-parent-property-outlook.md)|
|[Session](outlookbarstorage-session-property-outlook.md)|

## See also


#### Other resources


[Outlook Object Model Reference](http://msdn.microsoft.com/library/73221b13-d8d8-99b8-3394-b95dbbfd5ddc%28Office.15%29.aspx)
