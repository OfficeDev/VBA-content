---
title: Trendlines Object (PowerPoint)
keywords: vbapp10.chm721000
f1_keywords:
- vbapp10.chm721000
ms.prod: powerpoint
api_name:
- PowerPoint.Trendlines
ms.assetid: 8ac46695-aae0-3611-ebf7-c7339ea733ab
ms.date: 06/08/2017
---


# Trendlines Object (PowerPoint)

Represents a collection of all the  **[Trendline](trendline-object-powerpoint.md)** objects for the specified series.


## Remarks

Each  **Trendline** object represents a trendline in a chart. A trendline shows the trend, or direction, of data in a series.


## Example




 **Note**  Although the following code applies to Microsoft Word, you can readily modify it to apply to PowerPoint.

Use the  **[Trendlines](series-trendlines-method-powerpoint.md)** method to return the **Trendlines** collection. The following example displays the number of trendlines for series one of the first chart in the active document.




```vb
With ActiveDocument.InlineShapes(1)

    If .HasChart Then

        MsgBox .Chart.SeriesCollection(1).Trendlines.Count

    End If

End With
```




 **Note**  Although the following code applies to Microsoft Word, you can readily modify it to apply to PowerPoint.

Use the  **[Add](trendlines-add-method-powerpoint.md)** method to create a new trendline and add it to the series. The following example adds a linear trendline to the first series for the first chart in the active document.




```vb
With ActiveDocument.InlineShapes(1)
    If .HasChart Then
        .Chart.SeriesCollection(1) _
            .Trendlines.Add Type:=xlLinear, Name:="Linear Trend"
    End If
End With
```




 **Note**  Although the following code applies to Microsoft Word, you can readily modify it to apply to PowerPoint.

Use  **[Trendlines](series-trendlines-method-powerpoint.md)** (Index), where Index is the trendline index number, to return a single **TrendLine** object. The following example changes the trendline type for the first series of the first chart in the active document. If the series has no trendline, this example will fail.

The index number denotes the order in which the trendlines were added to the series.  `Trendlines(1)` is the first trendline added to the series, and `Trendlines(Trendlines.Count)` is the last one added.




```vb
With ActiveDocument.InlineShapes(1)

    If .HasChart Then

        .Chart.SeriesCollection(1).Trendlines(1).Type = xlMovingAvg

    End If

End With
```


## See also


#### Concepts


[PowerPoint Object Model Reference](object-model-powerpoint-vba-reference.md)

