---
title: Chart.SeriesCollection Method (PowerPoint)
keywords: vbapp10.chm684043
f1_keywords:
- vbapp10.chm684043
ms.prod: powerpoint
api_name:
- PowerPoint.Chart.SeriesCollection
ms.assetid: 8adeb8b4-ba4f-6cdf-33bf-dceb1845dfb8
ms.date: 06/08/2017
---


# Chart.SeriesCollection Method (PowerPoint)

Returns all the series in the chart.


## Syntax

 _expression_. **SeriesCollection**( **_Index_** )

 _expression_ A variable that represents a **[Chart](chart-object-powerpoint.md)** object.


### Return Value

A  **[SeriesCollection](seriescollection-object-powerpoint.md)** object that represents all the series in the chart.


## Example




 **Note**  Although the following code applies to Microsoft Word, you can readily modify it to apply to PowerPoint.

The following example turns on data labels for series one of the first chart in the active document.




```vb
With ActiveDocument.InlineShapes(1)

    If .HasChart Then

        .Chart.SeriesCollection(1).HasDataLabels = True

    End If

End With


```


## See also


#### Concepts


[Chart Object](chart-object-powerpoint.md)

