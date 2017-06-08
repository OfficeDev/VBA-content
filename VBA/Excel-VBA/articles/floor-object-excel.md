---
title: Floor Object (Excel)
keywords: vbaxl10.chm611072
f1_keywords:
- vbaxl10.chm611072
ms.prod: excel
api_name:
- Excel.Floor
ms.assetid: 74c71ca8-a0d4-f7cf-a002-5cec7a27b70d
ms.date: 06/08/2017
---


# Floor Object (Excel)

Represents the floor of a 3-D chart.


## Example

Use the  **[Floor](chart-floor-property-excel.md)** property to return the **Floor** object. The following example sets the floor color for embedded chart one to cyan. The example will fail if the chart isn't a 3-D chart.


```vb
Worksheets("sheet1").ChartObjects(1).Activate 
ActiveChart.Floor.Interior.Color = RGB(0, 255, 255)
```


## See also


#### Other resources


[Excel Object Model Reference](http://msdn.microsoft.com/library/11ea8598-8a20-92d5-f98b-0da04263bf2c%28Office.15%29.aspx)


