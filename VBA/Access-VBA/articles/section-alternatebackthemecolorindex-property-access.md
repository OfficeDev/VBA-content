---
title: Section.AlternateBackThemeColorIndex Property (Access)
keywords: vbaac10.chm14607
f1_keywords:
- vbaac10.chm14607
ms.prod: access
api_name:
- Access.Section.AlternateBackThemeColorIndex
ms.assetid: 15ef17dd-06fd-db4a-7253-5d193f2e4b9a
ms.date: 06/08/2017
---


# Section.AlternateBackThemeColorIndex Property (Access)

Gets or sets a value that represents a color in the applied color theme associated with the  **AlternateBackColor** property of the section. Read/write **Long**.


## Syntax

 _expression_. **AlternateBackThemeColorIndex**

 _expression_ A variable that represents a **Section** object.


## Remarks

The  **AlternateBackThemeColorIndex** property uses one of the values listed in the following table.



|**Value**|**Description**|
|:-----|:-----|
|0 |Text 1|
|1 |Background 1|
|2|Text 2|
|3 (Default)|Background 2|
|4|Accent 1|
|5|Accent 2|
|6|Accent 3|
|7|Accent 4|
|8|Accent 5|
|9|Accent 6|
|10|Hyperlink|
|11|Followed Hyperlink|
If no theme is applied, the  **AlternateBackThemeColorIndex** property contains -1.

This property is not surfaced in the property sheet.


## See also


#### Concepts


[Section Object](section-object-access.md)

