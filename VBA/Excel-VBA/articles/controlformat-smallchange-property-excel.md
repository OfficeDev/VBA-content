---
title: ControlFormat.SmallChange Property (Excel)
keywords: vbaxl10.chm630089
f1_keywords:
- vbaxl10.chm630089
ms.prod: excel
api_name:
- Excel.ControlFormat.SmallChange
ms.assetid: 5c2c668a-3d4d-ac01-e08b-0db6278ddffd
ms.date: 06/08/2017
---


# ControlFormat.SmallChange Property (Excel)

Returns or sets the amount that the scroll bar or spinner is incremented or decremented for a line scroll (when the user clicks an arrow). Read/write  **Long** .


## Syntax

 _expression_ . **SmallChange**

 _expression_ A variable that represents a **ControlFormat** object.


## Example

This example creates a scroll bar and sets its linked cell, minimum, maximum, large change, and small change values.


```vb
Set sb = Worksheets(1).Shapes.AddFormControl(xlScrollBar, _ 
 Left:=10, Top:=10, Width:=10, Height:=200) 
With sb.ControlFormat 
 .LinkedCell = "D1" 
 .Max = 100 
 .Min = 0 
 .LargeChange = 10 
 .SmallChange = 2 
End With
```


## See also


#### Concepts


[ControlFormat Object](controlformat-object-excel.md)

