---
title: OutlookBarPane Object (Outlook)
keywords: vbaol11.chm3003
f1_keywords:
- vbaol11.chm3003
ms.prod: outlook
api_name:
- Outlook.OutlookBarPane
ms.assetid: f8e6aa05-7a66-64f2-5a6a-ea639b6bbc59
ms.date: 06/08/2017
---


# OutlookBarPane Object (Outlook)

Represents the  **Shortcuts** pane in an explorer window.


## Remarks

Use the  **[Item](panes-item-method-outlook.md)** method to retrieve the **OutlookBarPane** object from a **[Panes](panes-object-outlook.md)** object. Because the **[Name](outlookbarpane-name-property-outlook.md)** property is the default property of the **OutlookBarPane** object, you can identify the **OutlookBarPane** object by name. For example:


## Example

The following example retrieves an  **OutlookBarPane** object by name.


```
Set myOlBarPane = myPanes.Item("OutlookBar")
```


## Events



|**Name**|
|:-----|
|[BeforeNavigate](outlookbarpane-beforenavigate-event-outlook.md)|

## Properties



|**Name**|
|:-----|
|[Application](outlookbarpane-application-property-outlook.md)|
|[Class](outlookbarpane-class-property-outlook.md)|
|[Contents](outlookbarpane-contents-property-outlook.md)|
|[Name](outlookbarpane-name-property-outlook.md)|
|[Parent](outlookbarpane-parent-property-outlook.md)|
|[Session](outlookbarpane-session-property-outlook.md)|
|[Visible](outlookbarpane-visible-property-outlook.md)|

## See also


#### Other resources


[Outlook Object Model Reference](http://msdn.microsoft.com/library/73221b13-d8d8-99b8-3394-b95dbbfd5ddc%28Office.15%29.aspx)
