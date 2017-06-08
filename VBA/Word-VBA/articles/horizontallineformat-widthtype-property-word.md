---
title: HorizontalLineFormat.WidthType Property (Word)
keywords: vbawd10.chm165543941
f1_keywords:
- vbawd10.chm165543941
ms.prod: word
api_name:
- Word.HorizontalLineFormat.WidthType
ms.assetid: 41d2ecc7-625e-5a62-8f68-f2982e04c6db
ms.date: 06/08/2017
---


# HorizontalLineFormat.WidthType Property (Word)

Returns or sets the width type for the specified  **HorizontalLineFormat** object. Read/write **WdHorizontalLineWidthType** .


## Syntax

 _expression_ . **WidthType**

 _expression_ Required. A variable that represents a **[HorizontalLineFormat](horizontallineformat-object-word.md)** object.


## Example

This example adds horizontal lines to the active document and compares their width types.


```vb
Dim temp As InlineShape 
Set temp = _ 
 ActiveDocument.InlineShapes.AddHorizontalLineStandard 
MsgBox "AddHorizontalLineStandard - WidthType = " _ 
 &; temp.HorizontalLineFormat.WidthType 
Set temp = _ 
 ActiveDocument.InlineShapes.AddHorizontalLine _ 
 ("C:\My Documents\ArtsyRule.gif") 
MsgBox "AddHorizontalLine - WidthType = " _ 
 &; temp.HorizontalLineFormat.WidthType
```


## See also


#### Concepts


[HorizontalLineFormat Object](horizontallineformat-object-word.md)

