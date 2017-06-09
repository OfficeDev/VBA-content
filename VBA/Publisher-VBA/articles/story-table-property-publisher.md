---
title: Story.Table Property (Publisher)
keywords: vbapb10.chm5832710
f1_keywords:
- vbapb10.chm5832710
ms.prod: publisher
api_name:
- Publisher.Story.Table
ms.assetid: e9da80d3-ea3c-b47c-d434-498c72955c14
ms.date: 06/08/2017
---


# Story.Table Property (Publisher)

Returns a  **Table** object that represents a table in Microsoft Publisher.


## Syntax

 _expression_. **Table**

 _expression_A variable that represents a  **Story** object.


## Example

The following example adds a 5x5 table on the first page of the active publication, and then selects the first column of the new table.


```vb
Sub NewTable() 
 With ActiveDocument.Pages(1).Shapes.AddTable(NumRows:=5, _ 
 NumColumns:=5, Left:=72, Top:=300, Width:=400, Height:=100) 
 .Table.Columns(3).Cells(3).Fill.ForeColor.RGB = RGB _ 
 (Red:=255, Green:=0, Blue:=0) 
 End With 
End Sub
```

The following example selects the specified table in the active publication. This example assumes that there is at least one shape on the first page of the active publication.




```vb
Sub SelectTable() 
 With ActiveDocument.Pages(1).Shapes(1) 
 If .Type = pbTable Then 
 .Table.Rows(3).Cells(3).Fill.ForeColor _ 
 .RGB = RGB(Red:=150, Green:=150, Blue:=150) 
 End If 
 End With 
End Sub
```


