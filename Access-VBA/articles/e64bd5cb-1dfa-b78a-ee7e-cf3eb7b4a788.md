
# VaryByCategories Property

 **Last modified:** July 28, 2015

 **True** if Microsoft Graph assigns a different color or pattern to each data marker. The chart must contain only one series. Read/write **Boolean**.

## Example

This example assigns a different color or pattern to each data marker in chart group one. The example should be run on a 2-D line chart that has data markers on a series.


```
myChart.ChartGroups(1).VaryByCategories = True
```

