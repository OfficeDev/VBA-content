---
title: Column.Width Property (Publisher)
keywords: vbapb10.chm4980739
f1_keywords:
- vbapb10.chm4980739
ms.prod: publisher
api_name:
- Publisher.Column.Width
ms.assetid: 9596b828-a5ce-e501-db59-a0e1533108b3
ms.date: 06/08/2017
---


# Column.Width Property (Publisher)

Returns or sets a  **Variant** that represents the width (in points) of a specified table column or shape. Read/write.


## Syntax

 _expression_. **Width**

 _expression_A variable that represents a  **Column** object.


## Example

This example creates a new table and sets the height and width of the second row and column, respectively.


```vb
Sub SetRowHeightColumnWidth() 
 With ActiveDocument.Pages(1).Shapes.AddTable(NumRows:=3, _ 
 NumColumns:=3, Left:=80, Top:=80, Width:=400, Height:=12).Table 
 .Rows(2).Height = 72 
 .Columns(2).Width = 72 
 End With 
End Sub
```


