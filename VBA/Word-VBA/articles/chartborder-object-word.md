---
title: ChartBorder Object (Word)
keywords: vbawd10.chm931
f1_keywords:
- vbawd10.chm931
ms.prod: word
api_name:
- Word.ChartBorder
ms.assetid: eea90670-c599-2ec8-5b7b-c946a4bcd638
ms.date: 06/08/2017
---


# ChartBorder Object (Word)

Represents the border of an object.


## Remarks

Most bordered objects have a border that is treated as a single entity, regardless of how many sides it has. The entire border must be returned as a unit. To return a  **Border** object, use the **Border** property for the particular bordered object (for example, the **[Border](trendline-border-property-word.md)** property of a **[TrendLine](trendline-object-word.md)** object).


## Example

 The following example changes the type and line style of a trendline on the active chart.


```vb
With ActiveDocument.InlineShapes(1).Chart.SeriesCollection(1).Trendlines(1) 
 .Type = xlLinear 
 .Border.LineStyle = xlDash 
End With
```


## See also


#### Other resources


[Word Object Model Reference](http://msdn.microsoft.com/library/be452561-b436-bb9b-6f94-3faa9a74a6fd%28Office.15%29.aspx)


