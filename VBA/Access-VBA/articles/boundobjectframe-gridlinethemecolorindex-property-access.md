---
title: BoundObjectFrame.GridlineThemeColorIndex Property (Access)
keywords: vbaac10.chm14635
f1_keywords:
- vbaac10.chm14635
ms.prod: access
api_name:
- Access.BoundObjectFrame.GridlineThemeColorIndex
ms.assetid: 1d527006-46f3-fc31-a579-ff2b32a104cc
ms.date: 06/08/2017
---


# BoundObjectFrame.GridlineThemeColorIndex Property (Access)

Gets or sets the theme color index that represents a color in the applied color theme associated with the  **GridlineColor** property of the specified object. Read/write **Long**.


## Syntax

 _expression_. **GridlineThemeColorIndex**

 _expression_ A variable that represents a **BoundObjectFrame** object.


## Remarks

The  **GridlineThemeColorIndex** property uses one of the values listed in the following table.



|**Value**|**Description**|
|:-----|:-----|
|0 |Text 1|
|1 (Default)|Background 1|
|2|Text 2|
|3|Background 2|
|4|Accent 1|
|5|Accent 2|
|6|Accent 3|
|7|Accent 4|
|8|Accent 5|
|9|Accent 6|
|10|Hyperlink|
|11|Followed Hyperlink|
If no theme is applied, the  **GridlineThemeColorIndex** property contains -1.

This property is not surfaced in the property sheet.


## See also


#### Concepts


[BoundObjectFrame Object](boundobjectframe-object-access.md)

