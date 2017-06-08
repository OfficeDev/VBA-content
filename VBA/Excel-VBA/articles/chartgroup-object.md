---
title: ChartGroup Object
keywords: vbagr10.chm131097
f1_keywords:
- vbagr10.chm131097
ms.prod: excel
api_name:
- Excel.ChartGroup
ms.assetid: 8a485a8c-e181-a039-60b9-a02c2c89b26e
ms.date: 06/08/2017
---


# ChartGroup Object

Represents one or more series of points plotted in a chart with the same format. A chart contains one or more chart groups, each chart group contains one or more  [series](series-object.md), and each series contains one or more  [points](point-object.md). For example, a single chart might contain both a line chart group, which contains all the series plotted with the line chart format, and a bar chart group, which contains all the series plotted with the bar chart format. The  **ChartGroup** object is a member of the **[ChartGroups](chartgroups-collection-excel.md)** collection.


## Using the ChartGroup Object

Use  **ChartGroups**( _index_), where  _index_ is the chart group's index number, to return a single **ChartGroup** object. The following example adds drop lines to chart group one in the chart.


```vb
myChart.ChartGroups(1).HasDropLines = True
```

Because the index number for a particular chart group can change if the chart format used for that group is changed, it may be easier to use one of the named shortcut methods for chart groups to return a particular chart group. The  **PieGroups** method returns the collection of pie chart groups in a chart, the **LineGroups** method returns the collection of all the line chart groups, and so on. You can use each of these methods with an index number to return a single **ChartGroup** object, or you can use each one without an index number to return a **ChartGroups** collection. The following methods are available for chart groups:


-  **[AreaGroups](areagroups-method.md)** method
    
-  **[BarGroups](bargroups-method.md)** method
    
-  **[ColumnGroups](columngroups-method.md)** method
    
-  **[DoughnutGroups](doughnutgroups-method.md)** method
    
-  **[LineGroups](linegroups-method.md)** method
    
-  **[PieGroups](piegroups-method.md)** method
    

