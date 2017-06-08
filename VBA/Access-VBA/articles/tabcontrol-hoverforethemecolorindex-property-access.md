---
title: TabControl.HoverForeThemeColorIndex Property (Access)
keywords: vbaac10.chm14616
f1_keywords:
- vbaac10.chm14616
ms.prod: access
api_name:
- Access.TabControl.HoverForeThemeColorIndex
ms.assetid: 75ffbe8c-2257-8cb7-4ee0-ab6e22e52a4f
ms.date: 06/08/2017
---


# TabControl.HoverForeThemeColorIndex Property (Access)

Gets or sets the theme color index that represents a color in the applied color theme associated with the  **HoverForeColor** property of the specified object. Read/write **Long**.


## Syntax

 _expression_. **HoverForeThemeColorIndex**

 _expression_ A variable that represents a **TabControl** object.


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


[TabControl Object](tabcontrol-object-access.md)

