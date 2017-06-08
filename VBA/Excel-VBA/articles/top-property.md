---
title: Top Property
keywords: vbagr10.chm65662
f1_keywords:
- vbagr10.chm65662
ms.prod: excel
api_name:
- Excel.Top
ms.assetid: 57938f4c-cd1f-b420-154d-fe4a8775c826
ms.date: 06/08/2017
---


# Top Property

Top property as it applies to the  **Application** object.

Returns or sets the position of the Application object. The distance from the top edge of the screen to the top edge of the main Microsoft Graph window. In Windows, if the application window is minimized, this property controls the position of the window icon (anywhere on the screen). Read/write Double.

 _expression_. **Top**

 _expression_ Required. An expression that returns one of the above objects.
Top property as it applies to the  **AxisTitle**,  **ChartArea**,  **ChartTitle**,  **DataLabel**,  **DataSheet**,  **DisplayUnitLabel**,  **Legend**, and  **PlotArea** objects.
The distance from the top edge of the object to the top of row 1 (on a datasheet) or the top of the chart area (on a chart). Read/write Double.
 _expression_. **Top**
 _expression_ Required. An expression that returns one of the above objects.
Top property as it applies to the  **Axis**,  **LegendEntry**, and  **LegendKey** objects.
The distance from the top edge of the object to the top of row 1 (on a datasheet) or the top of the chart area (on a chart). Read-only Double.
 _expression_. **Top**
 _expression_ Required. An expression that returns one of the above objects.
Top property as it applies to the  **Chart** object.
The distance from the top edge of the object to the top of row 1 (on a datasheet) or the top of the chart area (on a chart). Read/write Variant.
 _expression_. **Top**
 _expression_ Required. An expression that returns one of the above objects.

## Example

This example sets the position of the top of the chart title.


```
myChart.ChartTitle.Top = 10
```


