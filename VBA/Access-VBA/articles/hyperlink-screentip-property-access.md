---
title: Hyperlink.ScreenTip Property (Access)
keywords: vbaac10.chm10119
f1_keywords:
- vbaac10.chm10119
ms.prod: access
api_name:
- Access.Hyperlink.ScreenTip
ms.assetid: b935ea5c-17d8-e3ad-fca2-ef0985daa709
ms.date: 06/08/2017
---


# Hyperlink.ScreenTip Property (Access)

You can use the  **ScreenTip** property to specify or determine the text that is displayed when you move the cursor over a hyperlink control. Read/write **String**.


## Syntax

 _expression_. **ScreenTip**

 _expression_ A variable that represents a **Hyperlink** object.


## Remarks

When you move the cursor over a hyperlink control whose  **HyperlinkSubAddress** property is set, Microsoft Access changes the cursor to an upward-pointing hand and displays the text string defined by the **ScreenTip** property. Clicking the control displays the object or Web page specified by the link.

For more information about hyperlink addresses and their format, see the  **HyperlinkAddress** and **HyperlinkSubAddress** property topics.


## Example

The following example displays the message "Go to Home page" when the cursor hovers over the hyperlink named "HomePage" on the "Order Entry" form.


```vb
Forms("Order Entry").Controls("HomePage").Hyperlink.ScreenTip = "Go to Home page"
```


## See also


#### Concepts


[Hyperlink Object](hyperlink-object-access.md)

