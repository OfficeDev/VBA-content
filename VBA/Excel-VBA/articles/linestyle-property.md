---
title: LineStyle Property
keywords: vbagr10.chm3077078
f1_keywords:
- vbagr10.chm3077078
ms.prod: excel
api_name:
- Excel.LineStyle
ms.assetid: 4783a76a-9e73-c605-ade5-be8fec821b1d
ms.date: 06/08/2017
---


# LineStyle Property

Returns or sets the line style for the border. Read/write 
 **XlLineStyle**
.



|XlLineStyle can be one of these XlLineStyle constants.|
| **xlContinuous**|
| **xlDash**|
| **xlDashDot** **xlDashDotDot** **xlDot** **xlDouble** **xlSlantDashDot** **xlLineStyleNone**|

 _expression_. **LineStyle**

 _expression_ Required. An expression that returns one of the objects in the Applies To list.

## Example

This example puts a border around the chart area and the plot area.


```vb
With myChart 
 .ChartArea.Border.LineStyle = xlDashDot 
 With .PlotArea.Border 
 .LineStyle = xlDashDotDot 
 .Weight = xlThick 
 End With 
End With
```


