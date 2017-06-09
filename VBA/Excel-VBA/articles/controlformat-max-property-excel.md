---
title: ControlFormat.Max Property (Excel)
keywords: vbaxl10.chm630085
f1_keywords:
- vbaxl10.chm630085
ms.prod: excel
api_name:
- Excel.ControlFormat.Max
ms.assetid: 35ed65e1-94d7-c147-2535-d41c503bb19b
ms.date: 06/08/2017
---


# ControlFormat.Max Property (Excel)

Returns or sets the maximum value of a scroll bar or spinner range. The scroll bar or spinner won't take on values greater than this maximum value. Read/write  **Long** .


## Syntax

 _expression_ . **Max**

 _expression_ An expression that returns a **ControlFormat** object.


### Return Value

Long


## Remarks

The value of the  **Max** property must be greater than the value of the **[Min](controlformat-min-property-excel.md)** property.


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

