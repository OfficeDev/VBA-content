---
title: Point Object (PowerPoint)
keywords: vbapp10.chm714000
f1_keywords:
- vbapp10.chm714000
ms.prod: powerpoint
api_name:
- PowerPoint.Point
ms.assetid: e0137fdd-5632-88d7-a6c0-57a76717e736
ms.date: 06/08/2017
---


# Point Object (PowerPoint)

Represents a single point in a series in a chart.


## Remarks

 The **Point** object is a member of the **[Points](points-object-powerpoint.md)** collection. The **Points** collection contains all the points in one series.


## Example




 **Note**  Although the following code applies to Microsoft Word, you can readily modify it to apply to PowerPoint.

Use  **[Points](series-points-method-powerpoint.md)** ( _Index_ ), where _Index_ is the point index number, to return a single **Point** object. Points are numbered from left to right on the series. `Points(1)` is the leftmost point, and `Points(Points.Count)` is the rightmost point. The following example sets the marker style for the third point in series one for the first chart in the active document. The specified series must be a 2-D line, scatter, or radar series.




```vb
With ActiveDocument.InlineShapes(1)

    If .HasChart Then

        .Chart.SeriesCollection(1).Points(3).MarkerStyle = xlDiamond

    End If

End With


```


## See also


#### Concepts


[PowerPoint Object Model Reference](object-model-powerpoint-vba-reference.md)

