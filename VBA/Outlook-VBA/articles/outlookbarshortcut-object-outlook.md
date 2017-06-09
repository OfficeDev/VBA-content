---
title: OutlookBarShortcut Object (Outlook)
keywords: vbaol11.chm337
f1_keywords:
- vbaol11.chm337
ms.prod: outlook
api_name:
- Outlook.OutlookBarShortcut
ms.assetid: fae05770-1b06-1ddd-e2db-8428e64bd1e2
ms.date: 06/08/2017
---


# OutlookBarShortcut Object (Outlook)

Represents a shortcut in a group in the  **Shortcuts** pane.


## Remarks

Use the  **[Item](outlookbarshortcuts-item-method-outlook.md)** method to retrieve the **OutlookBarShortcut** object from an **[OutlookBarShortcuts](outlookbarshortcuts-object-outlook.md)** object. Because the **[Name](outlookbarshortcut-name-property-outlook.md)** property is the default property of the **OutlookBarShortcut** object, you can identify the shortcut by name.


## Example

The following example retrieves an  **OutlookBarShortcut** object by name.


```
Set myOlBarShortcut = myOutlookBarShortcuts.Item("Calendar")
```


## Methods



|**Name**|
|:-----|
|[SetIcon](outlookbarshortcut-seticon-method-outlook.md)|

## Properties



|**Name**|
|:-----|
|[Application](outlookbarshortcut-application-property-outlook.md)|
|[Class](outlookbarshortcut-class-property-outlook.md)|
|[Name](outlookbarshortcut-name-property-outlook.md)|
|[Parent](outlookbarshortcut-parent-property-outlook.md)|
|[Session](outlookbarshortcut-session-property-outlook.md)|
|[Target](outlookbarshortcut-target-property-outlook.md)|

## See also


#### Other resources


[Outlook Object Model Reference](http://msdn.microsoft.com/library/73221b13-d8d8-99b8-3394-b95dbbfd5ddc%28Office.15%29.aspx)
