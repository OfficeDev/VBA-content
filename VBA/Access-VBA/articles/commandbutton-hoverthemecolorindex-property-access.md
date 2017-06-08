---
title: CommandButton.HoverThemeColorIndex Property (Access)
keywords: vbaac10.chm14612
f1_keywords:
- vbaac10.chm14612
ms.prod: access
api_name:
- Access.CommandButton.HoverThemeColorIndex
ms.assetid: 7fec39e2-f79f-1260-ff6f-9e634ff18fe0
ms.date: 06/08/2017
---


# CommandButton.HoverThemeColorIndex Property (Access)

Gets or sets the theme color index that represents a color in the applied color theme associated with the  **HoverColor** property of the specified object. Read/write **Long**.


## Syntax

 _expression_. **HoverThemeColorIndex**

 _expression_ A variable that represents a **CommandButton** object.


## Remarks

The  **HoverThemeColorIndex** property uses one of the values listed in the following table.



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

