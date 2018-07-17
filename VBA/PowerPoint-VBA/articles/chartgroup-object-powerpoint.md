---
title: ChartGroup Object (PowerPoint)
keywords: vbapp10.chm692000
f1_keywords:
- vbapp10.chm692000
ms.prod: powerpoint
api_name:
- PowerPoint.ChartGroup
ms.assetid: 5caa5855-bd69-3fbc-f601-504e431a42e9
ms.date: 06/08/2017
---


# ChartGroup Object (PowerPoint)

Represents one or more series plotted in a chart with the same format.


## Remarks

A chart contains one or more chart groups, each chart group contains one or more **[Series](series-object-powerpoint.md)** objects, and each series contains one or more **[Points](points-object-powerpoint.md)** objects. For example, a single chart might contain both a line chart group, which contains all the series plotted with the line chart format, and a bar chart group, which contains all the series plotted with the bar chart format. The **ChartGroup** object is a member of the **[ChartGroups](chartgroups-object-powerpoint.md)** collection.

Use  **ChartGroups** ( _Index_ ), where _index_ is the chart group index number, to return a single **ChartGroup** object.


## Example




 **Note**  Although the following code applies to Microsoft Word, you can readily modify it to apply to PowerPoint.

The following example adds drop lines to the first chart group of the first chart in the active document.




```vb
With ActiveDocument.InlineShapes(1).Chart

    .ChartGroups(1).HasDropLines = True

End With
```


## See also


#### Concepts


[PowerPoint Object Model Reference](object-model-powerpoint-vba-reference.md)

