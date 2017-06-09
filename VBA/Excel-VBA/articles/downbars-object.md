---
title: DownBars Object
keywords: vbagr10.chm5207323
f1_keywords:
- vbagr10.chm5207323
ms.prod: excel
api_name:
- Excel.DownBars
ms.assetid: d85f4fac-c708-efe1-88c5-c2dca6616f31
ms.date: 06/08/2017
---


# DownBars Object

Represents the down bars in the specified chart group. Down bars connect points in the first series in the chart group with lower values in the last series (the lines go down from the first series). Only 2-D line groups that contain at least two series can have down bars. This object isn't a collection. There's no object that represents a single down bar; either you have up bars and down bars turned on for all points in a chart group or you have them turned off.


## Using the DownBars Object

Use the  **DownBars** property to return the **DownBars** object. The following example turns on up and down bars for chart group one in the chart. The example then sets the up-bar color to blue and the down-bar color to red.


```vb
With myChart.ChartGroups(1) 
 .HasUpDownBars = True 
 .UpBars.Interior.Color = RGB(0, 0, 255) 
 .DownBars.Interior.Color = RGB(255, 0, 0) 
End With
```


## Remarks

If the  **[HasUpDownBars](hasupdownbars-property.md)** property is  **False**, most properties of the  **DownBars** object are disabled.


