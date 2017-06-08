---
title: Shading.BackgroundPatternColor Property (Word)
keywords: vbawd10.chm154796037
f1_keywords:
- vbawd10.chm154796037
ms.prod: word
api_name:
- Word.Shading.BackgroundPatternColor
ms.assetid: 0d78f926-0fe6-aa37-bd39-c7233a5bf3e8
ms.date: 06/08/2017
---


# Shading.BackgroundPatternColor Property (Word)

Returns or sets the 24-bit color that's applied to the background of the  **Shading** object. Read/write.


## Syntax

 _expression_ . **BackgroundPatternColor**

 _expression_ Required. A variable that represents a **[Shading](shading-object-word.md)** object.


## Remarks

This property can be any valid  **WdColor** constant or a value returned by Visual Basic's **RGB** function.


## Example

This example applies turquoise background shading to the first paragraph in the active document.


```vb
Set myRange = ActiveDocument.Paragraphs(1).Range 
myRange.Shading.BackgroundPatternColor = _ 
 wdColorTurquoise
```

This example adds a table at the insertion point and then applies light gray background shading to the first cell.




```vb
Selection.Collapse Direction:=wdCollapseStart 
Set myTable = _ 
 ActiveDocument.Tables.Add(Range:=Selection.Range, _ 
 NumRows:=2, NumColumns:=2) 
myTable.Cell(1, 1).Shading.BackgroundPatternColor = _ 
 wdColorGray25
```


## See also


#### Concepts


[Shading Object](shading-object-word.md)

