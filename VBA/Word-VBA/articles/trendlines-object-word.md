---
title: Trendlines Object (Word)
keywords: vbawd10.chm1562
f1_keywords:
- vbawd10.chm1562
ms.prod: word
api_name:
- Word.Trendlines
ms.assetid: 06c20a75-4afc-03f5-1eec-eee1559d3f52
ms.date: 06/08/2017
---


# Trendlines Object (Word)

Represents a collection of all the  **[Trendline](trendline-object-word.md)** objects for the specified series.


## Remarks

Each  **Trendline** object represents a trendline in a chart. A trendline shows the trend, or direction, of data in a series.


## Example

Use the  **[Trendlines](series-trendlines-method-word.md)** method to return the **Trendlines** collection. The following example displays the number of trendlines for series one of the first chart in the active document.


```vb
With ActiveDocument.InlineShapes(1) 
 If .HasChart Then 
 MsgBox .Chart.SeriesCollection(1).Trendlines.Count 
 End If 
End With
```

Use the  **[Add](trendlines-add-method-word.md)** method to create a new trendline and add it to the series. The following example adds a linear trendline to the first series for the first chart in the active document.




```vb
With ActiveDocument.InlineShapes(1) 
 If .HasChart Then 
 .Chart.SeriesCollection(1) _ 
 .Trendlines.Add Type:=xlLinear, Name:="Linear Trend" 
 End If 
End With
```

Use  **[Trendlines](series-trendlines-method-word.md)** (Index), where Index is the trendline index number, to return a single **TrendLine** object. The following example changes the trendline type for the first series of the first chart in the active document. If the series has no trendline, this example will fail.

The index number denotes the order in which the trendlines were added to the series.  `Trendlines(1)` is the first trendline added to the series, and `Trendlines(Trendlines.Count)` is the last one added.




```vb
With ActiveDocument.InlineShapes(1) 
 If .HasChart Then 
 .Chart.SeriesCollection(1).Trendlines(1).Type = xlMovingAvg 
 End If 
End With
```


## See also


#### Other resources


[Word Object Model Reference](http://msdn.microsoft.com/library/be452561-b436-bb9b-6f94-3faa9a74a6fd%28Office.15%29.aspx)


