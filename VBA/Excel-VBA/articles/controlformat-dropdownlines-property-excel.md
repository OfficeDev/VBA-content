---
title: ControlFormat.DropDownLines Property (Excel)
keywords: vbaxl10.chm630076
f1_keywords:
- vbaxl10.chm630076
ms.prod: excel
api_name:
- Excel.ControlFormat.DropDownLines
ms.assetid: e2e12163-c247-6518-2d2f-701d27266a1c
ms.date: 06/08/2017
---


# ControlFormat.DropDownLines Property (Excel)

Returns or sets the number of list lines displayed in the drop-down portion of a combo box. Read/write  **Long** .


## Syntax

 _expression_ . **DropDownLines**

 _expression_ A variable that represents a **ControlFormat** object.


## Example

This example creates a combo box with 10 list lines.


```vb
With Worksheets(1).Shapes.AddFormControl(xlDropDown, _ 
 Left:=10, Top:=10, Width:=100, Height:=10) 
 .ControlFormat.DropDownLines = 10 
End With
```


## See also


#### Concepts


[ControlFormat Object](controlformat-object-excel.md)

