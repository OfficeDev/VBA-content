---
title: Points Object (Word)
ms.prod: word
api_name:
- Word.Points
ms.assetid: d0adc45a-7b31-a25e-d96f-f2a098702501
ms.date: 06/08/2017
---


# Points Object (Word)

A collection of all the  **[Point](point-object-word.md)** objects in the specified series in a chart.


## Remarks

Use  **[Points](series-points-method-word.md)** ( _Index_ ), where _Index_ is the point index number, to return a single **Point** object. Points are numbered from left to right on the series. `Points(1)` is the leftmost point, and `Points(Points.Count)` is the rightmost point.


## Example

Use the  **Points** method to return the **Points** collection. The following example adds a data label to the last point in series one for the first chart in the active document.


```vb
With ActiveDocument.InlineShapes(1) 
 If .HasChart Then 
 With .Chart.SeriesCollection(1).Points 
 .Item(.Count).ApplyDataLabels Type:=xlShowValue 
 End With 
 End If 
End With
```

 The following example sets the marker style for the third point in series one for the first chart in the active document. The specified series must be a 2-D line, scatter, or radar series.




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


