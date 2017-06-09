---
title: Shading.BackgroundPatternColorIndex Property (Word)
keywords: vbawd10.chm154796034
f1_keywords:
- vbawd10.chm154796034
ms.prod: word
api_name:
- Word.Shading.BackgroundPatternColorIndex
ms.assetid: 47e78b6a-4519-3b8a-9d26-39ead1019d43
ms.date: 06/08/2017
---


# Shading.BackgroundPatternColorIndex Property (Word)

Returns or sets the color that's applied to the background of the  **Shading** object. Read/write **WdColorIndex** .


## Syntax

 _expression_ . **BackgroundPatternColorIndex**

 _expression_ Required. A variable that represents a **[Shading](shading-object-word.md)** object.


## Example

This example applies cyan background shading to the first paragraph in the active document.


```vb
Dim rngTemp As Range 
 
Set rngTemp = ActiveDocument.Paragraphs(1).Range 
rngTemp.Shading.BackgroundPatternColorIndex = wdTurquoise
```

This example adds a table at the insertion point and then applies light gray background shading to the first cell.




```vb
Dim tableNew As Table 
 
Selection.Collapse Direction:=wdCollapseStart 
Set tableNew = ActiveDocument.Tables.Add(Range:=Selection.Range, _ 
 NumRows:=2, NumColumns:=2) 
tableNew.Cell(1, 1).Shading.BackgroundPatternColorIndex = _ 
 wdGray25
```


## See also


#### Concepts


[Shading Object](shading-object-word.md)

