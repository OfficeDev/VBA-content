---
title: Point.SecondaryPlot Property (PowerPoint)
keywords: vbapp10.chm67198
f1_keywords:
- vbapp10.chm67198
ms.prod: powerpoint
api_name:
- PowerPoint.Point.SecondaryPlot
ms.assetid: 37bba3d7-2bb7-fd46-eaf8-eb8b44aa071c
ms.date: 06/08/2017
---


# Point.SecondaryPlot Property (PowerPoint)

 **True** if the point is in the secondary section of either a pie-of-pie chart or a bar-of-pie chart. Read/write **Boolean**.


## Syntax

 _expression_. **SecondaryPlot**

 _expression_ A variable that represents a **[Point](point-object-powerpoint.md)** object.


## Remarks

This property applies only to points on pie-of-pie charts or bar-of-pie charts. 


## Example




 **Note**  Although the following code applies to Microsoft Word, you can readily modify it to apply to PowerPoint.

The following example moves point four to the secondary section of the chart. You must run this example on either a pie-of-pie chart or a bar-of-pie chart. 




```vb
With ActiveDocument.InlineShapes(1)

    If .HasChart Then

        With .Chart.SeriesCollection(1)

            .Points(4).SecondaryPlot = True

        End With

    End If

End With
```


## See also


#### Concepts


[Point Object](point-object-powerpoint.md)

