---
title: ToggleButton.PressedForeThemeColorIndex Property (Access)
keywords: vbaac10.chm14624
f1_keywords:
- vbaac10.chm14624
ms.prod: access
api_name:
- Access.ToggleButton.PressedForeThemeColorIndex
ms.assetid: 9c2b6020-3bb5-72f5-184d-2b1453946a26
ms.date: 06/08/2017
---


# ToggleButton.PressedForeThemeColorIndex Property (Access)

Gets or sets the theme color index that represents a color in the applied color theme associated with the  **PressedForeColor** property of the specified object. Read/write **Long**.


## Syntax

 _expression_. **PressedForeThemeColorIndex**

 _expression_ A variable that represents a **ToggleButton** object.


## Remarks

The  **PressedForeThemeColorIndex** uses one of the values listed in the following table.



|**Value**|**Description**|
|:-----|:-----|
|0|Text 1|
|1 |Background 1|
|2|Text 2|
|3|Background 2|
|4|Accent 1|
|5|Accent 2|
|6|Accent 3|
|7|Accent 4|
|8 (Default)|Accent 5|
|9|Accent 6|
|10|Hyperlink|
|11|Followed Hyperlink|
If no theme is applied, the  **PressedForeThemeColorIndex** property contains -1.

This property is not surfaced in the property sheet.


## See also


#### Concepts


[ToggleButton Object](togglebutton-object-access.md)

