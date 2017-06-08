---
title: DropLines Object (Excel)
keywords: vbaxl10.chm603072
f1_keywords:
- vbaxl10.chm603072
ms.prod: excel
api_name:
- Excel.DropLines
ms.assetid: 88fdf5f5-2842-2d68-a073-18d05fd2fa38
ms.date: 06/08/2017
---


# DropLines Object (Excel)

Represents the drop lines in a chart group.


## Remarks

Drop lines connect the points in the chart with the x-axis. Only line and area chart groups can have drop lines. This object isn't a collection. There's no object that represents a single drop line; you either have drop lines turned on for all points in a chart group or you have them turned off.

If the  **[HasDropLines](chartgroup-hasdroplines-property-excel.md)** property is **False** , most properties of the **DropLines** object are disabled.


## Example

Use the  **DropLines** property to return the **DropLines** object. The following example turns on drop lines for chart group one in embedded chart one and then sets the drop line color to red.


```vb
Worksheets("sheet1").ChartObjects(1).Activate 
ActiveChart.ChartGroups(1).HasDropLines = True 
ActiveChart.ChartGroups(1).DropLines.Border.ColorIndex = 3
```


## See also


#### Other resources


[Excel Object Model Reference](http://msdn.microsoft.com/library/11ea8598-8a20-92d5-f98b-0da04263bf2c%28Office.15%29.aspx)


