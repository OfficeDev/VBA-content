---
title: DownBars Object (Excel)
keywords: vbaxl10.chm609072
f1_keywords:
- vbaxl10.chm609072
ms.prod: excel
api_name:
- Excel.DownBars
ms.assetid: 23623e02-44c7-a6b2-e3a8-fffc4f7b3164
ms.date: 06/08/2017
---


# DownBars Object (Excel)

Represents the down bars in a chart group.


## Remarks

 Down bars connect points on the first series in the chart group with lower values on the last series (the lines go down from the first series). Only 2-D line groups that contain at least two series can have down bars. This object isn't a collection. There's no object that represents a single down bar; you either have up bars and down bars turned on for all points in a chart group or you have them turned off.

If the  **[HasUpDownBars](chartgroup-hasupdownbars-property-excel.md)** property is **False** , most properties of the **DownBars** object are disabled.


## Example

Use the  **[DownBars](chartgroup-downbars-property-excel.md)** property to return the **DownBars** object. The following example turns on up and down bars for chart group one in embedded chart one on the worksheet named "Sheet5." The example then sets the up bar color to blue and the down bar color to red.


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


