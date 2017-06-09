---
title: HiLoLines Object
keywords: vbagr10.chm5207532
f1_keywords:
- vbagr10.chm5207532
ms.prod: excel
api_name:
- Excel.HiLoLines
ms.assetid: 6793025e-0b3e-360c-4292-02397395535a
ms.date: 06/08/2017
---


# HiLoLines Object

Represents the high-low lines in the specified chart group. High-low lines connect the highest point with the lowest point in every category in the chart group. Only 2-D line groups can have high-low lines. This object isn't a collection. There's no object that represents a single high-low line; either you have high-low lines turned on for all points in a chart group or you have them turned off.


## Using the HiLoLines Object

Use the  **HiLoLines** property to return the **HiLoLines** object. The following example makes the high-low lines in chart group one in the chart blue.


```
myChart.ChartGroups(1).HiLoLines.Border.Color = RGB(0, 0, 255)
```


## Remarks

If the  **[HasHiLoLines](hashilolines-property.md)** property is  **False**, most properties of the  **HiLoLines** object are disabled.


