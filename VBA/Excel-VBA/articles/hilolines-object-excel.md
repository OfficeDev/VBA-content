---
title: HiLoLines Object (Excel)
keywords: vbaxl10.chm599072
f1_keywords:
- vbaxl10.chm599072
ms.prod: excel
api_name:
- Excel.HiLoLines
ms.assetid: 3248f878-4be9-acbd-3515-70f8255b4d69
ms.date: 06/08/2017
---


# HiLoLines Object (Excel)

Represents the high-low lines in a chart group.


## Remarks

 High-low lines connect the highest point with the lowest point in every category in the chart group. Only 2-D line groups can have high-low lines. This object isn't a collection. There's no object that represents a single high-low line; you either have high-low lines turned on for all points in a chart group or you have them turned off.

If the  **[HasHiLoLines](chartgroup-hashilolines-property-excel.md)** property is **False** , most properties of the **HiLoLines** object are disabled.


## Example

Use the  **[HiLoLines](chartgroup-hilolines-property-excel.md)** property to return the **HiLoLines** object. The following example uses the **HasHiLowLines** property to add HiLowLines to embedded chart one (the chart must be a line chart) on worksheet one. The example then makes the high-low lines blue.


```vb
Worksheets(1).ChartObjects(1).Activate 
ActiveChart.ChartGroups(1).HasHiLoLines = True 
ActiveChart.ChartGroups(1).HiLoLines.Border.Color = RGB(0, 0, 255)
```


## See also


#### Other resources


[Excel Object Model Reference](http://msdn.microsoft.com/library/11ea8598-8a20-92d5-f98b-0da04263bf2c%28Office.15%29.aspx)


