---
title: Chart.RightAngleAxes Property (PowerPoint)
keywords: vbapp10.chm684040
f1_keywords:
- vbapp10.chm684040
ms.prod: powerpoint
api_name:
- PowerPoint.Chart.RightAngleAxes
ms.assetid: 4bccf442-1cf6-48b9-d67c-5a72561211e0
ms.date: 06/08/2017
---


# Chart.RightAngleAxes Property (PowerPoint)

 **True** if the chart axes are at right angles, independent of chart rotation or elevation. Read/write **Boolean**.


## Syntax

 _expression_. **RightAngleAxes**

 _expression_ A variable that represents a **[Chart](chart-object-powerpoint.md)** object.


## Remarks

This property applies only to 3-D line, column, and bar charts. 

If this property is set to  **True**, the **[Perspective](chart-perspective-property-powerpoint.md)** property is ignored.


## Example




 **Note**  Although the following code applies to Microsoft Word, you can readily modify it to apply to PowerPoint.

The following example sets the axes for the first chart in the active document to intersect at right angles. You should run the example on a 3-D chart.




```vb
With ActiveDocument.InlineShapes(1)

    If .HasChart Then

        .Chart.RightAngleAxes = True

    End If

End With
```


## See also


#### Concepts


[Chart Object](chart-object-powerpoint.md)

