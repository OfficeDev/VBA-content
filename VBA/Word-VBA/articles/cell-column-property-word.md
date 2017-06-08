---
title: Cell.Column Property (Word)
keywords: vbawd10.chm156106853
f1_keywords:
- vbawd10.chm156106853
ms.prod: word
api_name:
- Word.Cell.Column
ms.assetid: b3f5f0a1-4d17-9d66-f689-9eb6308132fe
ms.date: 06/08/2017
---


# Cell.Column Property (Word)

Returns a  **Column** object that represents the table column containing the specified cell. Read-only.


## Syntax

 _expression_ . **Column**

 _expression_ Required. A variable that represents a **[Cell](cell-object-word.md)** object.


## Example

This example creates a 3x5 table and applies shading to the even-numbered columns.


```vb
Dim tableNew As Table 
Dim cellLoop As Cell 
 
Selection.Collapse Direction:=wdCollapseStart 
Set tableNew = _ 
 ActiveDocument.Tables.Add(Range:=Selection.Range, _ 
 NumRows:=3, NumColumns:=5) 
 
For Each cellLoop In tableNew.Rows(1).Cells 
 If cellLoop.ColumnIndex Mod 2 = 0 Then 
 cellLoop.Column.Shading.Texture = wdTexture10Percent 
 End If 
Next cellLoop
```


## See also


#### Concepts


[Cell Object](cell-object-word.md)

