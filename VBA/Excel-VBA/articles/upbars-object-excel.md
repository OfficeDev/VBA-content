---
title: UpBars Object (Excel)
keywords: vbaxl10.chm607072
f1_keywords:
- vbaxl10.chm607072
ms.prod: excel
api_name:
- Excel.UpBars
ms.assetid: 4f2a85fe-3fbb-ccc6-7b16-e48e54cd3394
ms.date: 06/08/2017
---


# UpBars Object (Excel)

Represents the up bars in a chart group.


## Remarks

Up bars connect points on series one with higher values on the last series in the chart group (the lines go up from series one). Only 2-D line groups that contain at least two series can have up bars. This object isn't a collection. There's no object that represents a single up bar; you either have up bars turned on for all points in a chart group or you have them turned off.

If the  **[HasUpDownBars](chartgroup-hasupdownbars-property-excel.md)** property is **False** , most properties of the **UpBars** object are disabled.


## Example

Use the  **[UpBars](chartgroup-upbars-property-excel.md)** property to return the **UpBars** object. The following example turns on up and down bars for chart group one in embedded chart one on Sheet5. The example then sets the up bar color to blue and sets the down bar color to red.


```vb
With Worksheets("sheet5").ChartObjects(1).Chart.ChartGroups(1) 
 .HasUpDownBars = True 
 .UpBars.Interior.Color = RGB(0, 0, 255) 
 .DownBars.Interior.Color = RGB(255, 0, 0) 
End With
```


## See also


#### Other resources


[Excel Object Model Reference](http://msdn.microsoft.com/library/11ea8598-8a20-92d5-f98b-0da04263bf2c%28Office.15%29.aspx)


