---
title: Shape.ControlFormat Property (Excel)
keywords: vbaxl10.chm636128
f1_keywords:
- vbaxl10.chm636128
ms.prod: excel
api_name:
- Excel.Shape.ControlFormat
ms.assetid: e874098f-ea8c-93ff-f746-a0d568bec5b5
ms.date: 06/08/2017
---


# Shape.ControlFormat Property (Excel)

Returns a  **[ControlFormat](controlformat-object-excel.md)** object that contains Microsoft Excel control properties. Read-only.


## Syntax

 _expression_ . **ControlFormat**

 _expression_ A variable that represents a **Shape** object.


## Example

This example removes the selected item from a list box. If  `Shapes(2)` doesn't represent a list box, this example fails.


```vb
Set lbcf = Worksheets(1).Shapes(2).ControlFormat 
lbcf.RemoveItem lbcf.ListIndex
```


## See also


#### Concepts


[Shape Object](shape-object-excel.md)

