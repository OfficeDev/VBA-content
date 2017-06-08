---
title: CommandButton.ThemeFontIndex Property (Access)
keywords: vbaac10.chm14610
f1_keywords:
- vbaac10.chm14610
ms.prod: access
api_name:
- Access.CommandButton.ThemeFontIndex
ms.assetid: 8cb51c03-09a1-83ba-c6bf-7e74d7444c8b
ms.date: 06/08/2017
---


# CommandButton.ThemeFontIndex Property (Access)

Gets or sets the font index that represents a font in the applied theme associated with the  **FontName** property of the specified object. Read/write **Long**.


## Syntax

 _expression_. **ThemeFontIndex**

 _expression_ A variable that represents a **CommandButton** object.


## Remarks

The  **ThemeFontIndex** property uses one of the values listed in the following table.



|**Value**|**Description**|
|:-----|:-----|
|0|Header font|
|1 (Default)|Detail font|
If no theme is applied, the  **ThemeFontIndex** property contains -1.

This property is not surfaced in the property sheet.


## See also


#### Concepts


[CommandButton Object](commandbutton-object-access.md)

