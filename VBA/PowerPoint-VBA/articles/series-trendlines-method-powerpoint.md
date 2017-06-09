---
title: Series.Trendlines Method (PowerPoint)
keywords: vbapp10.chm65690
f1_keywords:
- vbapp10.chm65690
ms.prod: powerpoint
api_name:
- PowerPoint.Series.Trendlines
ms.assetid: 17578607-d0aa-dcc2-1eec-3af031f17c2d
ms.date: 06/08/2017
---


# Series.Trendlines Method (PowerPoint)

Returns a collection of all the trendlines for the series.


## Syntax

 _expression_. **Trendlines**( **_Index_** )

 _expression_ A variable that represents a **[Series](series-object-powerpoint.md)** object.


### Return Value

A  **[Trendlines](trendlines-object-powerpoint.md)** object that represents all the treadlines for the series.


## Example




 **Note**  Although the following code applies to Microsoft Word, you can readily modify it to apply to PowerPoint.

The following example adds a linear trendline to series one for the first chart in the active document.




```vb
With ActiveDocument.InlineShapes(1)

    If .HasChart Then

        .Chart.SeriesCollection(1).Trendlines.Add Type:=xlLinear

    End If

End With
```


## See also


#### Concepts


[Series Object](series-object-powerpoint.md)

