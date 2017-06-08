---
title: Axis.DisplayUnitLabel Property (Excel)
keywords: vbaxl10.chm561116
f1_keywords:
- vbaxl10.chm561116
ms.prod: excel
api_name:
- Excel.Axis.DisplayUnitLabel
ms.assetid: e3a78e7b-464e-80b0-8bde-49f08ab4c842
ms.date: 06/08/2017
---


# Axis.DisplayUnitLabel Property (Excel)

Returns the  **[DisplayUnitLabel](displayunitlabel-object-excel.md)** object for the specified axis. Returns **null** if the **[HasDisplayUnitLabel](axis-hasdisplayunitlabel-property-excel.md)** property is set to **False** . Read-only.


## Syntax

 _expression_ . **DisplayUnitLabel**

 _expression_ A variable that represents an **Axis** object.


## Example

This example sets the label caption to "Millions" for the value axis in Chart1, and then it turns off automatic font scaling.


```vb
With Charts("Chart1").Axes(xlValue).DisplayUnitLabel 
 .Caption = "Millions" 
 .AutoScaleFont = False 
End With
```


## See also


#### Concepts


[Axis Object](axis-object-excel.md)

