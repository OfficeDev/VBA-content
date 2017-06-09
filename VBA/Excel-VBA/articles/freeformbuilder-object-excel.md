---
title: FreeformBuilder Object (Excel)
keywords: vbaxl10.chm647072
f1_keywords:
- vbaxl10.chm647072
ms.prod: excel
api_name:
- Excel.FreeformBuilder
ms.assetid: 91c779ac-69bc-3b68-8ecb-1f9cc8e5b20e
ms.date: 06/08/2017
---


# FreeformBuilder Object (Excel)

Represents the geometry of a freeform while it's being built.


## Remarks

Use the  **[BuildFreeform](shapes-buildfreeform-method-excel.md)** method to return a **FreeformBuilder** object. Use the **[AddNodes](freeformbuilder-addnodes-method-excel.md)** method to add nodes to the freefrom. Use the **[ConvertToShape](freeformbuilder-converttoshape-method-excel.md)** method to create the shape defined in the **FreeformBuilder** object and add it to the **[Shapes](shapes-object-excel.md)** collection.


## Example

 The following example adds a freeform with four segments to _myDocument_ .


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


#### Other resources



[Excel Object Model Reference](http://msdn.microsoft.com/library/11ea8598-8a20-92d5-f98b-0da04263bf2c%28Office.15%29.aspx)

