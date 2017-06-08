---
title: CommandButton.HoverForeThemeColorIndex Property (Access)
keywords: vbaac10.chm14616
f1_keywords:
- vbaac10.chm14616
ms.prod: access
api_name:
- Access.CommandButton.HoverForeThemeColorIndex
ms.assetid: 7952f076-a8ac-c6d3-72f7-23e8365d8a16
ms.date: 06/08/2017
---


# CommandButton.HoverForeThemeColorIndex Property (Access)

Gets or sets the theme color index that represents a color in the applied color theme associated with the  **HoverForeColor** property of the specified object. Read/write **Long**.


## Syntax

 _expression_. **HoverForeThemeColorIndex**

 _expression_ A variable that represents a **CommandButton** object.


## Remarks

The  **HoverForeThemeColorIndex** property uses one of the values listed in the following table.



|**Value**|**Description**|
|:-----|:-----|
|0|Text 1|
|1 |Background 1|
|2|Text 2|
|3|Background 2|
|4|Accent 1|
|5|Accent 2|
|6|Accent 3|
|7 (Default)|Accent 4|
|8|Accent 5|
|9|Accent 6|
|10|Hyperlink|
|11|Followed Hyperlink|
If no theme is applied, the  **HoverThemeColorIndex** property contains -1.

This property is not surfaced in the property sheet.


## See also


#### Concepts


[CommandButton Object](commandbutton-object-access.md)

