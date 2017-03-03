---
title: ListBox.ThemeFontIndex Property (Access)
keywords: vbaac10.chm14610
f1_keywords:
- vbaac10.chm14610
ms.prod: ACCESS
api_name:
- Access.ListBox.ThemeFontIndex
ms.assetid: 7fa3a5ef-c59b-8ce5-1d7f-6b00991dc12b
---


# ListBox.ThemeFontIndex Property (Access)

Gets or sets the font index that represents a font in the applied theme associated with the  **FontName** property of the specified object. Read/write **Long**.


## Syntax

 _expression_. **ThemeFontIndex**

 _expression_ A variable that represents a **ListBox** object.


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


[ListBox Object](listbox-object-access.md)

