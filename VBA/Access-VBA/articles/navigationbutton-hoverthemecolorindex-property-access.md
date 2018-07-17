---
title: NavigationButton.HoverThemeColorIndex Property (Access)
keywords: vbaac10.chm14612
f1_keywords:
- vbaac10.chm14612
ms.prod: access
api_name:
- Access.NavigationButton.HoverThemeColorIndex
ms.assetid: d2412f7c-0344-1a32-e95b-5abd04d9aabe
ms.date: 06/08/2017
---


# NavigationButton.HoverThemeColorIndex Property (Access)

Gets or sets the theme color index that represents a color in the applied color theme associated with the  **HoverColor** property of the specified object. Read/write **Long**.


## Syntax

 _expression_. **HoverThemeColorIndex**

 _expression_ A variable that represents a **NavigationButton** object.


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


[NavigationButton Object](navigationbutton-object-access.md)

