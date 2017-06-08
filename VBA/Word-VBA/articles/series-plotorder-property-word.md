---
title: Series.PlotOrder Property (Word)
keywords: vbawd10.chm123732196
f1_keywords:
- vbawd10.chm123732196
ms.prod: word
api_name:
- Word.Series.PlotOrder
ms.assetid: 8813c546-f5ed-774e-e57f-3adfcb6ac926
ms.date: 06/08/2017
---


# Series.PlotOrder Property (Word)

Returns or sets the plot order for the selected series within the chart group. Read/write  **Long** .


## Syntax

 _expression_ . **PlotOrder**

 _expression_ A variable that represents a **[Chart](chart-object-word.md)** object.


## Remarks

You can set plot order only within a chart group (you cannot set the plot order for the entire chart if you have more than one chart type). A chart group is a collection of series with the same chart type.

Changing the plot order of one series will cause the plot orders of the other series in the chart group to be adjusted, as necessary.


## Example

The following example makes series two for the first chart in the active document appear third in the plot order. You should run the example on a 2-D column chart that contains three or more series.


```vb
With ActiveDocument.InlineShapes(1) 
 If .HasChart Then 
 .Chart.ChartGroups(1).SeriesCollection(2).PlotOrder = 3 
 End If 
End With
```


## See also


#### Concepts


[Series Object](series-object-word.md)

