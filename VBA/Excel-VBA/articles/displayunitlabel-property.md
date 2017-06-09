---
title: DisplayUnitLabel Property
keywords: vbagr10.chm67318
f1_keywords:
- vbagr10.chm67318
ms.prod: excel
api_name:
- Excel.DisplayUnitLabel
ms.assetid: 50e91894-9b5d-c915-e94c-e4563b54487a
ms.date: 06/08/2017
---


# DisplayUnitLabel Property

Returns the  **[DisplayUnitLabel](displayunitlabel-object.md)** object for the value axis in the specified chart. Returns  **Null** if the **[HasDisplayUnitLabel](hasdisplayunitlabel-property.md)** property is  **False**. Read-only.


## Example

This example sets the caption for the value axis in myChart to "Millions" and turns off automatic font scaling.


```vb
With myChart.Axes(xlValue).DisplayUnitLabel 
 .Caption = "Millions" 
 .AutoScaleFont = False 
End With
```


