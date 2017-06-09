---
title: Axis.HasDisplayUnitLabel Property (Word)
keywords: vbawd10.chm113049675
f1_keywords:
- vbawd10.chm113049675
ms.prod: word
api_name:
- Word.Axis.HasDisplayUnitLabel
ms.assetid: 0d5f02d5-241d-691b-4505-1eda392d6feb
ms.date: 06/08/2017
---


# Axis.HasDisplayUnitLabel Property (Word)

 **True** if the label specified by the **[DisplayUnit](axis-displayunit-property-word.md)** or **[DisplayUnitCustom](axis-displayunitcustom-property-word.md)** property is displayed on the specified axis. The default is **True** . Read/write **Boolean** .


## Syntax

 _expression_ . **HasDisplayUnitLabel**

 _expression_ A variable that represents an **[Axis](axis-object-word.md)** object.


## Example

The following example sets the units on the value axis of the first chart in the active document to increments of 500 but keeps the unit label hidden.


```vb
With ActiveDocument.InlineShapes(1) 
 If .HasChart Then 
 With .Chart.Axes(xlValue) 
 .DisplayUnit = xlCustom 
 .DisplayUnitCustom = 500 
 .AxisTitle.Caption = "Rebate Amounts" 
 .HasDisplayUnitLabel = False 
 End With 
 End If 
End With 

```


## See also


#### Concepts


[Axis Object](axis-object-word.md)

