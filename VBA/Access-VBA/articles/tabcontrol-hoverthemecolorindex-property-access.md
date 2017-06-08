---
title: TabControl.HoverThemeColorIndex Property (Access)
keywords: vbaac10.chm14612
f1_keywords:
- vbaac10.chm14612
ms.prod: access
api_name:
- Access.TabControl.HoverThemeColorIndex
ms.assetid: 9e8e2111-33b5-0dc8-5949-f6512b7603e4
ms.date: 06/08/2017
---


# TabControl.HoverThemeColorIndex Property (Access)

Gets or sets the theme color index that represents a color in the applied color theme associated with the  **HoverColor** property of the specified object. Read/write **Long**.


## Syntax

 _expression_. **HoverThemeColorIndex**

 _expression_ A variable that represents a **TabControl** object.


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


[TabControl Object](tabcontrol-object-access.md)

