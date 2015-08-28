
# ChartGroups Collection (Excel)

 **Last modified:** July 28, 2015

A collection of all the  ** [ChartGroup](8a485a8c-e181-a039-60b9-a02c2c89b26e.md)**objects in the specified chart. Each  **ChartGroup** object represents one or more series plotted with the same format in a chart. A chart contains one or more chart groups, each chart group contains one or more series, and each series contains one or more points. For example, a single chart might contain both a line chart group, containing all the series plotted with the line chart format, and a bar chart group, containing all the series plotted with the bar chart format.

## Using the ChartGroups Collection

Use the  **ChartGroups** method to return the **ChartGroups** collection. The following example displays the number of chart groups in the chart


```
MsgBox myChart.ChartGroups.Count
```

Use  **ChartGroups**( _index_), where  _index_ is the chart group's index number, to return a single **ChartGroup** object. The following example adds drop lines to chart group one in the chart.




```
myChart.ChartGroups(1).HasDropLines = True
```

Because the index number for a particular chart group can change if the chart format used for that group is changed, it may be easier to use one of the named chart-group shortcut methods to return a particular chart group. The  **PieGroups** method returns the collection of pie chart groups in the specified chart, the **LineGroups** method returns the collection of line chart groups, and so on. Each of these methods can be used with an index number to return a single **ChartGroup** object, or used without an index number to return a **ChartGroups** collection. The following methods are available for chart groups:


-  ** [AreaGroups](ec2a4a28-2f10-4f4f-bd91-642bf1b8ebe2.md)**method
    
-  ** [BarGroups](a00e484e-05ec-2eaa-cc33-05b77a4af0b5.md)**method
    
-  ** [ColumnGroups](dcb4d7e0-ce56-46d9-35d9-d9653bbb6f97.md)**method
    
-  ** [DoughnutGroups](41ca4213-c17b-7bba-c357-7ba65fd55d39.md)**method
    
-  ** [LineGroups](3a8083b5-8b71-e28b-c775-6be50544d6b2.md)**method
    
-  ** [PieGroups](f7fd5497-f7a0-6c28-1a59-9e6f37a0885e.md)**method
    
