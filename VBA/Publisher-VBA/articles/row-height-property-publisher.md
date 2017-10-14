---
title: Row.Height Property (Publisher)
keywords: vbapb10.chm4849667
f1_keywords:
- vbapb10.chm4849667
ms.prod: publisher
api_name:
- Publisher.Row.Height
ms.assetid: fd243edc-1521-bd67-a365-2c4685ee5037
ms.date: 06/08/2017
---


# Row.Height Property (Publisher)

Returns or sets a  **Variant** that represents the height (in points) of a specified table row or shape. Read/write.


## Syntax

 _expression_. **Height**

 _expression_A variable that represents a  **Row** object.


## Remarks

The valid range for the  **Height** property depends on the size of the application workspace and the position of the object within the workspace. For centered objects on non-banner page sizes, the **Height** property may be 0.0 to 50.0 inches. For centered objects on banner page sizes, the **Height** property may be 0.0 to 241.0 inches.


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


