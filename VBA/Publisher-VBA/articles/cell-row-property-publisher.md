---
title: Cell.Row Property (Publisher)
keywords: vbapb10.chm5111831
f1_keywords:
- vbapb10.chm5111831
ms.prod: publisher
api_name:
- Publisher.Cell.Row
ms.assetid: b961af2b-6b03-f54b-922e-d2e7633a3dfe
ms.date: 06/08/2017
---


# Cell.Row Property (Publisher)

Returns a  **Long** that represents the row number containing the specified cell. Read-only.


## Syntax

 _expression_. **Row**

 _expression_A variable that represents a  **Cell** object.


## Example

This example enters the fill for all even-numbered rows and clears the fill for all odd-numbered rows in the specified table. This example assumes the specified shape is a table and not another type of shape.


```vb
Sub FillCellsByRow() 
 Dim shpTable As Shape 
 Dim rowTable As Row 
 Dim celTable As Cell 
 
 Set shpTable = ActiveDocument.Pages(1).Shapes _ 
 .AddTable(NumRows:=5, NumColumns:=5, Left:=100, _ 
 Top:=100, Width:=100, Height:=12) 
 For Each rowTable In shpTable.Table.Rows 
 For Each celTable In rowTable.Cells 
 If celTable.Row Mod 2 = 0 Then 
 celTable.Fill.ForeColor.RGB = RGB _ 
 (Red:=180, Green:=180, Blue:=180) 
 Else 
 celTable.Fill.ForeColor.RGB = RGB _ 
 (Red:=255, Green:=255, Blue:=255) 
 End If 
 Next celTable 
 Next rowTable 
 
End Sub
```


