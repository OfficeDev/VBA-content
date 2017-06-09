---
title: Point Object (Word)
keywords: vbawd10.chm4000
f1_keywords:
- vbawd10.chm4000
ms.prod: word
api_name:
- Word.Point
ms.assetid: 349ea9a3-9e9a-b26f-146f-799d39c3d4a9
ms.date: 06/08/2017
---


# Point Object (Word)

Represents a single point in a series in a chart.


## Remarks

 The **Point** object is a member of the **[Points](points-object-word.md)** collection. The **Points** collection contains all the points in one series.


## Example

Use  **[Points](series-points-method-word.md)** ( _Index_ ), where _Index_ is the point index number, to return a single **Point** object. Points are numbered from left to right on the series. `Points(1)` is the leftmost point, and `Points(Points.Count)` is the rightmost point. The following example sets the marker style for the third point in series one for the first chart in the active document. The specified series must be a 2-D line, scatter, or radar series.


```vb
With ActiveDocument.InlineShapes(1) 
 If .HasChart Then 
 .Chart.SeriesCollection(1).Points(3).MarkerStyle = xlDiamond 
 End If 
End With 

```


## See also


#### Other resources


[Word Object Model Reference](http://msdn.microsoft.com/library/be452561-b436-bb9b-6f94-3faa9a74a6fd%28Office.15%29.aspx)


