---
title: Axis.Crosses Property (PowerPoint)
keywords: vbapp10.chm682005
f1_keywords:
- vbapp10.chm682005
ms.prod: powerpoint
api_name:
- PowerPoint.Axis.Crosses
ms.assetid: 93390bc6-8d94-4bf3-257e-c20fce2a2c62
ms.date: 06/08/2017
---


# Axis.Crosses Property (PowerPoint)

Returns or sets the point on the specified axis where the other axis crosses. Read/write  **Long**.


## Syntax

 _expression_. **Crosses**

 _expression_ A variable that represents an **[Axis](axis-object-powerpoint.md)** object.


## Remarks

This property is not available for radar charts. For 3-D charts, this property can be applied only to the value axis and indicates where the plane defined by the category axes crosses the value axis.

You can use this property for both category and value axes. On the category axis,  **xlMinimum** sets the value axis to cross at the first category, and **xlMaximum** sets the value axis to cross at the last category.

Note that  **xlMinimum** and **xlMaximum** can have different meanings, depending on the axis.


## Example




 **Note**  Although the following code applies to Microsoft Word, you can readily modify it to apply to PowerPoint.

The following example sets the value axis in for the first chart in the active document to cross the category axis at the maximum x-axis value.




```vb
With ActiveDocument.InlineShapes(1)

     If .HasChart Then

        .Chart.Axes(xlCategory).Crosses = xlMaximum

    End If

End With
```


## See also


#### Concepts


[Axis Object](axis-object-powerpoint.md)

