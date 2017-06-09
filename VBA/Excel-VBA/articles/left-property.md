---
title: Left Property
keywords: vbagr10.chm65663
f1_keywords:
- vbagr10.chm65663
ms.prod: excel
api_name:
- Excel.Left
ms.assetid: 9d300adc-3d72-02d5-e39c-c40e21b7e8d5
ms.date: 06/08/2017
---


# Left Property

Left property as it applies to the  **Application**, and  **DataSheet** object.

Returns or sets the distance from the left edge of the screen to the left edge of the main Microsoft Graph window. Read/write Double.

 _expression_. **Left**

 _expression_ Required. An expression that returns one of the above objects.

## Remarks

If the window is maximized, Application.Left returns a negative number that varies based on the width of the window border. Setting Application.Left to 0 (zero) will make the window a tiny bit smaller than it would be if the application window were maximized. In other words, if Application.Left is 0, the left border of the main Microsoft Graph window will just barely be visible on the screen.

If the Microsoft Graph window is minimized, Application.Left controls the position of the window icon.


## Example

As it applies to the  **ChartTitle** object.

This example aligns the left edge of the chart title with the left edge of the chart area.




```
myChart.ChartTitle.Left = 0 

```

Left property as it applies to the  **AxisTitle**,  **ChartArea**,  **ChartTitle**,  **DataLabel**,  **DisplayUnitLabel**,  **Legend**, and  **PlotArea** objects.
Returns or sets the distance from the left edge of the object to the left edge of the chart area. Read/write Double.
 _expression_. **Left**
 _expression_ Required. An expression that returns one of the above objects.
Left property as it applies to the  **Axis**,  **LegendEntry**, and  **LegendKey** objects.
Returns or sets the distance from the left edge of the object to the left edge of the chart area. Read-only Double.
 _expression_. **Left**
 _expression_ Required. An expression that returns one of the above objects.
Left property as it applies to the  **Chart** object.
Returns or sets the distance from the left edge of the object to the left edge of the Microsoft Graph window. Read/write Variant.
 _expression_. **Left**
 _expression_ Required. An expression that returns a **Chart** object.

