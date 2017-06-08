---
title: Selection.Columns Property (Word)
keywords: vbawd10.chm158662958
f1_keywords:
- vbawd10.chm158662958
ms.prod: word
api_name:
- Word.Selection.Columns
ms.assetid: 444726a7-0bbe-8d66-b3ca-113af074e673
ms.date: 06/08/2017
---


# Selection.Columns Property (Word)

Returns a  **Columns** collection that represents all the table columns in a selection. Read-only.


## Syntax

 _expression_ . **Columns**

 _expression_ An expression that returns a **[Selection](selection-object-word.md)** object.


## Example

This example sets the width of the current column to 1 inch.


```vb
If Selection.Information(wdWithInTable) = True Then 
 Selection.Columns.SetWidth ColumnWidth:=InchesToPoints(1), _ 
 RulerStyle:=wdAdjustProportional 
End If
```


## See also


#### Concepts


[Selection Object](selection-object-word.md)

