---
title: DisplayUnitLabel Object
keywords: vbagr10.chm131087
f1_keywords:
- vbagr10.chm131087
ms.prod: excel
api_name:
- Excel.DisplayUnitLabel
ms.assetid: 1d8f0340-1760-295a-2c4e-92709d1deabc
ms.date: 06/08/2017
---


# DisplayUnitLabel Object

Represents a unit label on the value axis in the specified chart. Unit labels are useful for charting large values—for example, numbers in the millions or billions. You can make the chart more readable by using a single unit label instead of large numbers with strings of zeros next to the tick marks on the axis. This way, you need never have numbers of more than one or two digits by the tick marks.


## Using the DisplayUnitLabel Object

Use the  **[DisplayUnitLabel](displayunitlabel-property.md)** property to return the  **DisplayUnitLabel** object. The following example sets the caption for the value axis in myChart to "Millions" and turns off automatic font scaling.


```vb
With myChart.Axes(xlValue).DisplayUnitLabel 
 .Caption = "Millions" 
 .AutoScaleFont = False 
End With
```


