---
title: OutlookBarShortcuts Object (Outlook)
keywords: vbaol11.chm3004
f1_keywords:
- vbaol11.chm3004
ms.prod: outlook
api_name:
- Outlook.OutlookBarShortcuts
ms.assetid: 5ee9f085-d2fe-c949-9edc-ad073801ea77
ms.date: 06/08/2017
---


# OutlookBarShortcuts Object (Outlook)

Contains a set of  **[OutlookBarShortcut](outlookbarshortcut-object-outlook.md)** objects representing all shortcuts in a group in the **Shortcuts** pane.


## Remarks

Use the  **[Shortcuts](outlookbargroup-shortcuts-property-outlook.md)** property to return the **OutlookBarShortcuts** collection object from the **[OutlookBarGroup](outlookbargroup-object-outlook.md)** object.


## Example

The following Microsoft Visual Basic for Applications (VBA) example shows how to retrieve the  **OutlookBarShortcuts** object.


```
Set myShortcuts = myOutlookBarGroup.Shortcuts
```


## Events



|**Name**|
|:-----|
|[BeforeShortcutAdd](outlookbarshortcuts-beforeshortcutadd-event-outlook.md)|
|[BeforeShortcutRemove](outlookbarshortcuts-beforeshortcutremove-event-outlook.md)|
|[ShortcutAdd](outlookbarshortcuts-shortcutadd-event-outlook.md)|

## Methods



|**Name**|
|:-----|
|[Add](outlookbarshortcuts-add-method-outlook.md)|
|[Item](outlookbarshortcuts-item-method-outlook.md)|
|[Remove](outlookbarshortcuts-remove-method-outlook.md)|

## Properties



|**Name**|
|:-----|
|[Application](outlookbarshortcuts-application-property-outlook.md)|
|[Class](outlookbarshortcuts-class-property-outlook.md)|
|[Count](outlookbarshortcuts-count-property-outlook.md)|
|[Parent](outlookbarshortcuts-parent-property-outlook.md)|
|[Session](outlookbarshortcuts-session-property-outlook.md)|

## See also


#### Other resources


[Outlook Object Model Reference](http://msdn.microsoft.com/library/73221b13-d8d8-99b8-3394-b95dbbfd5ddc%28Office.15%29.aspx)
