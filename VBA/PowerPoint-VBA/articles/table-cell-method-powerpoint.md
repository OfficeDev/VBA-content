---
title: Table.Cell Method (PowerPoint)
keywords: vbapp10.chm622005
f1_keywords:
- vbapp10.chm622005
ms.prod: powerpoint
api_name:
- PowerPoint.Table.Cell
ms.assetid: 31a2908b-7a33-994d-860a-e01da62729e7
ms.date: 06/08/2017
---


# Table.Cell Method (PowerPoint)

Returns a  **[Cell](cell-object-powerpoint.md)** object that represents a cell in a table.


## Syntax

 _expression_. **Cell**( **_Row_**, **_Column_** )

 _expression_ A variable that represents a **Table** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Row_|Required|**Long**|The number of the row in the table to return. Can be an integer between 1 and the number of rows in the table.|
| _Column_|Required|**Long**|The number of the column in the table to return. Can be an integer between 1 and the number of columns in the table.|

### Return Value

Cell


## Example

This example creates a 3x3 table on a new slide in a new presentation and inserts text into the first cell of the table.


```vb
With Presentations.Add 
    With .Slides.AddSlide(1, ppLayoutBlank) 
        .Shapes.AddTable(3, 3).Select 
        .Shapes(1).Table.Cell(1, 1).Shape.TextFrame _ 
            .TextRange.Text = "Cell 1" 
    End With 
End With
```

This example sets the thickness of the bottom border of the cell in row 2, column 1 to two points.




```vb
ActivePresentation.Slides(2).Shapes(5).Table _ 
    .Cell(2, 1).Borders(ppBorderBottom).Weight = 2
```


## See also


#### Concepts


[Table Object](table-object-powerpoint.md)

