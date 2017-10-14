---
title: Table.Rows Property (PowerPoint)
keywords: vbapp10.chm622004
f1_keywords:
- vbapp10.chm622004
ms.prod: powerpoint
api_name:
- PowerPoint.Table.Rows
ms.assetid: f7003d61-62d4-8d00-15c5-d9a2c5d57625
ms.date: 06/08/2017
---


# Table.Rows Property (PowerPoint)

Returns a  **[Rows](rows-object-powerpoint.md)** collection that represents all the rows in a table. Read-only.


## Syntax

 _expression_. **Rows**

 _expression_ A variable that represents a **Table** object.


### Return Value

Rows


## Remarks

For information about returning a single member of a collection, see [Returning an Object from a Collection](return-objects-from-collections.md).


## Example

This example deletes the third row from the table in shape five of slide two in the active presentation.


```vb
ActivePresentation.Slides(2).Shapes(5).Table.Rows(3).Delete
```

This example applies a dashed line style to the bottom border of the second row of table cells.




```vb
ActiveWindow.Selection.ShapeRange.Table.Rows(2) _
    .Cells.Borders(ppBorderBottom).DashStyle = msoLineDash
```


## See also


#### Concepts


[Table Object](table-object-powerpoint.md)

