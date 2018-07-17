---
title: Cell.BorderLeft Property (Publisher)
keywords: vbapb10.chm5111812
f1_keywords:
- vbapb10.chm5111812
ms.prod: publisher
api_name:
- Publisher.Cell.BorderLeft
ms.assetid: f996a96f-4392-48c2-e5c2-bfe373a7997a
ms.date: 06/08/2017
---


# Cell.BorderLeft Property (Publisher)

Returns a  [CellBorder](cellborder-object-publisher.md)object that represents the left border for a specified table cell.


## Syntax

 _expression_. **BorderLeft**

 _expression_A variable that represents a  **Cell** object.


### Return Value

CellBorder


## Example

This example creates a checkerboard design using borders and a fill color with an existing table. This assumes the first shape on page two is a table and not another type of shape and that the table has an uneven number of columns.


```vb
Sub FillCellsByRow() 
 Dim shpTable As Shape 
 Dim rowTable As Row 
 Dim celTable As Cell 
 Dim intCell As Integer 
 
 intCell = 1 
 
 Set shpTable = ActiveDocument.Pages(2).Shapes(1) 
 For Each rowTable In shpTable.Table.Rows 
 For Each celTable In rowTable.Cells 
 With celTable 
 With .BorderBottom 
 .Weight = 2 
 .Color.RGB = RGB(Red:=0, Green:=0, Blue:=0) 
 End With 
 With .BorderTop 
 .Weight = 2 
 .Color.RGB = RGB(Red:=0, Green:=0, Blue:=0) 
 End With 
 With .BorderLeft 
 .Weight = 2 
 .Color.RGB = RGB(Red:=0, Green:=0, Blue:=0) 
 End With 
 With .BorderRight 
 .Weight = 2 
 .Color.RGB = RGB(Red:=0, Green:=0, Blue:=0) 
 End With 
 End With 
 If intCell Mod 2 = 0 Then 
 celTable.Fill.ForeColor.RGB = RGB _ 
 (Red:=180, Green:=180, Blue:=180) 
 Else 
 celTable.Fill.ForeColor.RGB = RGB _ 
 (Red:=255, Green:=255, Blue:=255) 
 End If 
 intCell = intCell + 1 
 Next celTable 
 Next rowTable 
 
End Sub
```


