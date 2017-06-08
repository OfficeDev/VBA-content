---
title: ChartGroup Object (Excel)
keywords: vbaxl10.chm567072
f1_keywords:
- vbaxl10.chm567072
ms.prod: excel
api_name:
- Excel.ChartGroup
ms.assetid: 7eee66c5-04a7-fd86-6e34-4c22ccaf8de0
ms.date: 06/08/2017
---


# ChartGroup Object (Excel)

Represents one or more series plotted in a chart with the same format.


## Remarks

A chart contains one or more chart groups, each chart group contains one or more **[Series](series-object-excel.md)** objects, and each series contains one or more **[Points](point-object-excel.md)** objects. For example, a single chart might contain both a line chart group, containing all the series plotted with the line chart format, and a bar chart group, containing all the series plotted with the bar chart format. The **ChartGroup** object is a member of the **[ChartGroups](chartgroups-object-excel.md)** collection.

Use  **ChartGroups** ( _index_ ), where _index_ is the chart-group index number, to return a single **ChartGroup** object.

Because the index number for a particular chart group can change if the chart format used for that group is changed, it may be easier to use one of the named chart group shortcut methods to return a particular chart group. The  **PieGroups** method returns the collection of pie chart groups in a chart, the **LineGroups** method returns the collection of line chart groups, and so on. Each of these methods can be used with an index number to return a single **ChartGroup** object, or without an index number to return a **ChartGroups** collection.


## Example

The following example adds drop lines to chart group 1 on chart sheet 1.


```vb
Charts(1).ChartGroups(1).HasDropLines = True
```

If the chart has been activated, you can use the  **ActiveChart** property.




```vb
Charts(1).Activate 
ActiveChart.ChartGroups(1).HasDropLines = True
```


## See also


#### Other resources



[Excel Object Model Reference](http://msdn.microsoft.com/library/11ea8598-8a20-92d5-f98b-0da04263bf2c%28Office.15%29.aspx)

