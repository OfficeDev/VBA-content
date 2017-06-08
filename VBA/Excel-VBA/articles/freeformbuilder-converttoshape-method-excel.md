---
title: FreeformBuilder.ConvertToShape Method (Excel)
keywords: vbaxl10.chm648074
f1_keywords:
- vbaxl10.chm648074
ms.prod: excel
api_name:
- Excel.FreeformBuilder.ConvertToShape
ms.assetid: 2084277d-7e6a-5675-8e46-17522c3228eb
ms.date: 06/08/2017
---


# FreeformBuilder.ConvertToShape Method (Excel)

Creates a shape that has the geometric characteristics of the specified  **[FreeformBuilder](freeformbuilder-object-excel.md)** object. Returns a **[Shape](shape-object-excel.md)** object that represents the new shape.


## Syntax

 _expression_ . **ConvertToShape**

 _expression_ A variable that represents a **FreeformBuilder** object.


### Return Value

Shape


## Remarks

 You must apply the **[AddNodes](freeformbuilder-addnodes-method-excel.md)** method to a **FreeformBuilder** object at least once before you use the **ConvertToShape** method.


## Example

This example adds a freeform with five vertices to  `myDocument`.


```vb
Set myDocument = Worksheets(1) 
With myDocument.Shapes.BuildFreeform(msoEditingCorner, 360, 200) 
 .AddNodes msoSegmentCurve, msoEditingCorner, _ 
 380, 230, 400, 250, 450, 300 
 .AddNodes msoSegmentCurve, msoEditingAuto, 480, 200 
 .AddNodes msoSegmentLine, msoEditingAuto, 480, 400 
 .AddNodes msoSegmentLine, msoEditingAuto, 360, 200 
 .ConvertToShape 
End With
```


## See also


#### Concepts


[FreeformBuilder Object](freeformbuilder-object-excel.md)

