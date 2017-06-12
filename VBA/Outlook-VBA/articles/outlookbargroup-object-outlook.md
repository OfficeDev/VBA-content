---
title: OutlookBarGroup Object (Outlook)
keywords: vbaol11.chm320
f1_keywords:
- vbaol11.chm320
ms.prod: outlook
api_name:
- Outlook.OutlookBarGroup
ms.assetid: 4ccc4213-5a57-7a8b-4ce5-869a096bd096
ms.date: 06/08/2017
---


# OutlookBarGroup Object (Outlook)

Represents a group of shortcuts in the  **Shortcuts** pane of an explorer window.


## Remarks

Use the  **[Item](outlookbargroups-item-method-outlook.md)** method to retrieve the **OutlookBarGroup** object from an **[OutlookBarGroups](outlookbargroups-object-outlook.md)** object. Because the **[Name](outlookbargroup-name-property-outlook.md)** property is the default property of the **OutlookBarGroup** object, you can identify the group by name.


## Example

The following example retrieves an  **OutlookBarGroup** object by name.


```
Set myOlBarGroup = myOutlookBarGroups.Item("Other Shortcuts")
```


## Properties



|**Name**|
|:-----|
|[Application](outlookbargroup-application-property-outlook.md)|
|[Class](outlookbargroup-class-property-outlook.md)|
|[Name](outlookbargroup-name-property-outlook.md)|
|[Parent](outlookbargroup-parent-property-outlook.md)|
|[Session](outlookbargroup-session-property-outlook.md)|
|[Shortcuts](outlookbargroup-shortcuts-property-outlook.md)|
|[ViewType](outlookbargroup-viewtype-property-outlook.md)|

## See also


#### Other resources


[Outlook Object Model Reference](http://msdn.microsoft.com/library/73221b13-d8d8-99b8-3394-b95dbbfd5ddc%28Office.15%29.aspx)
