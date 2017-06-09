---
title: ChartGroups Collection (Excel)
keywords: vbagr10.chm5207191
f1_keywords:
- vbagr10.chm5207191
ms.prod: excel
ms.assetid: 203bc32b-61e7-9bbc-bcc3-c7d8afc8b2ae
ms.date: 06/08/2017
---


# ChartGroups Collection (Excel)

A collection of all the  **[ChartGroup](chartgroup-object.md)** objects in the specified chart. Each  **ChartGroup** object represents one or more series plotted with the same format in a chart. A chart contains one or more chart groups, each chart group contains one or more series, and each series contains one or more points. For example, a single chart might contain both a line chart group, containing all the series plotted with the line chart format, and a bar chart group, containing all the series plotted with the bar chart format.


## Using the ChartGroups Collection

Use the  **ChartGroups** method to return the **ChartGroups** collection. The following example displays the number of chart groups in the chart


```vb
MsgBox myChart.ChartGroups.Count
```

Use  **ChartGroups**( _index_), where  _index_ is the chart group's index number, to return a single **ChartGroup** object. The following example adds drop lines to chart group one in the chart.




```vb
myChart.ChartGroups(1).HasDropLines = True
```

Because the index number for a particular chart group can change if the chart format used for that group is changed, it may be easier to use one of the named chart-group shortcut methods to return a particular chart group. The  **PieGroups** method returns the collection of pie chart groups in the specified chart, the **LineGroups** method returns the collection of line chart groups, and so on. Each of these methods can be used with an index number to return a single **ChartGroup** object, or used without an index number to return a **ChartGroups** collection. The following methods are available for chart groups:


-  **[AreaGroups](areagroups-method.md)** method
    
-  **[BarGroups](bargroups-method.md)** method
    
-  **[ColumnGroups](columngroups-method.md)** method
    
-  **[DoughnutGroups](doughnutgroups-method.md)** method
    
-  **[LineGroups](linegroups-method.md)** method
    
-  **[PieGroups](piegroups-method.md)** method
    

