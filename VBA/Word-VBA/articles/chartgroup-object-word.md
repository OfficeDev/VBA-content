---
title: ChartGroup Object (Word)
keywords: vbawd10.chm4020
f1_keywords:
- vbawd10.chm4020
ms.prod: word
api_name:
- Word.ChartGroup
ms.assetid: ea5a2610-9c00-9c95-8366-f9b0fcdf90be
ms.date: 06/08/2017
---


# ChartGroup Object (Word)

Represents one or more series plotted in a chart with the same format.


## Remarks

A chart contains one or more chart groups, each chart group contains one or more **[Series](series-object-word.md)** objects, and each series contains one or more **[Points](points-object-word.md)** objects. For example, a single chart might contain both a line chart group, which contains all the series plotted with the line chart format, and a bar chart group, which contains all the series plotted with the bar chart format. The **ChartGroup** object is a member of the **[ChartGroups](chartgroups-object-word.md)** collection.

Use  **ChartGroups** ( _Index_ ), where _index_ is the chart group index number, to return a single **ChartGroup** object.


## Example

The following example adds drop lines to the first chart group of the first chart in the active document.


```vb
With ActiveDocument.InlineShapes(1).Chart 
 .ChartGroups(1).HasDropLines = True 
End With
```


## See also


#### Other resources


[Word Object Model Reference](http://msdn.microsoft.com/library/be452561-b436-bb9b-6f94-3faa9a74a6fd%28Office.15%29.aspx)


