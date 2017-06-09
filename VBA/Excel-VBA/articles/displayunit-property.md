---
title: DisplayUnit Property
keywords: vbagr10.chm3077025
f1_keywords:
- vbagr10.chm3077025
ms.prod: excel
api_name:
- Excel.DisplayUnit
ms.assetid: c86b932e-6314-068f-f06e-4f35ead883d4
ms.date: 06/08/2017
---


# DisplayUnit Property

Returns or sets the units displayed for the value axis in the specified chart. If the value is xlCustom, the DisplayUnitCustom property returns or sets the value of the units displayed for the value axis. Read/write XlDisplayUnit .



|XlDisplayUnit can be one of these XlDisplayUnit constants.|
| **xlHundredMillions**|
| **xlHundreds**|
| **xlHundredThousands**|
| **xlMillionMillions**|
| **xlMillions**|
| **xlTenMillions**|
| **xlTenThousands**|
| **xlThousandMillions**|
| **xlThousands**|

 _expression_. **DisplayUnit**

 _expression_ Required. An expression that returns one of the objects in the Applies To list.

## Remarks

Using unit labels for the value axis when charting large values makes the incremental labels on the axis more readable and the data easier to track. In other words, if you label your value axis in thousands (for example), you can use smaller numeric values next to the tick marks on the axis.


## Example

This example sets the units displayed on the value axis in myChart to hundreds.


```vb
With myChart.Axes(xlValue) 
 .DisplayUnit = xlHundreds 
 .HasTitle = True 
 .AxisTitle.Caption = "Rebate Amounts" 
End With
```


