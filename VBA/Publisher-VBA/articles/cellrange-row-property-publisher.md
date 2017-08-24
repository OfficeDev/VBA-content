---
title: CellRange.Row Property (Publisher)
keywords: vbapb10.chm5177350
f1_keywords:
- vbapb10.chm5177350
ms.prod: publisher
api_name:
- Publisher.CellRange.Row
ms.assetid: ac5bccf0-6c9b-ce0e-20e5-f06ef29886c6
ms.date: 06/08/2017
---


# CellRange.Row Property (Publisher)

Returns a  **Long** that represents the row number containing the specified cell. Read-only.


## Syntax

 _expression_. **Row**

 _expression_A variable that represents a  **CellRange** object.


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


