---
title: Axis.AxisBetweenCategories Property (PowerPoint)
keywords: vbapp10.chm682001
f1_keywords:
- vbapp10.chm682001
ms.prod: powerpoint
api_name:
- PowerPoint.Axis.AxisBetweenCategories
ms.assetid: 8e0e0e80-58b9-005f-c719-ad45b491f9a9
ms.date: 06/08/2017
---


# Axis.AxisBetweenCategories Property (PowerPoint)

 **True** if the value axis crosses the category axis between categories. Read/write **Boolean**.


## Syntax

 _expression_. **AxisBetweenCategories**

 _expression_ A variable that represents an **[Axis](axis-object-powerpoint.md)** object.


## Remarks

This property applies only to category axes, and it does not apply to 3-D charts.


## Example




 **Note**  Although the following code applies to Microsoft Word, you can readily modify it to apply to PowerPoint.

The following example causes the value axis for the first chart in the active document to cross the category axis between categories.




```vb
With ActiveDocument.InlineShapes(1)
    If .HasChart Then
        .Chart.Axes(xlCategory). _
            AxisBetweenCategories = True
    End If
End With
```


## See also


#### Concepts


[Axis Object](axis-object-powerpoint.md)

