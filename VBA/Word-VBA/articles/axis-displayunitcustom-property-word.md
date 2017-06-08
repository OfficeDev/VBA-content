---
title: Axis.DisplayUnitCustom Property (Word)
keywords: vbawd10.chm113049673
f1_keywords:
- vbawd10.chm113049673
ms.prod: word
api_name:
- Word.Axis.DisplayUnitCustom
ms.assetid: 578e195b-9e45-1265-b20e-8de6a8233272
ms.date: 06/08/2017
---


# Axis.DisplayUnitCustom Property (Word)

If the value of the  **[DisplayUnit](axis-displayunit-property-word.md)** property is **xlCustom** , returns or sets the value of the displayed units. Read/write **Double** .


## Syntax

 _expression_ . **DisplayUnitCustom**

 _expression_ A variable that represents an **[Axis](axis-object-word.md)** object.


## Remarks

The value of this property must be from 0 through 10E307.

Using unit labels when charting large values makes your tick-mark labels easier to read. For example, if you label your value axis in units of hundreds, thousands, or millions, you can use smaller numeric values at the tick marks on the axis.


## Example

The following example sets the units displayed on the value axis of the first chart in the active document to increments of 500.


```vb
With ActiveDocument.InlineShapes(1) 
 If .HasChart Then 
 With .Chart.Axes(xlValue) 
 .DisplayUnit = xlCustom 
 .DisplayUnitCustom = 500 
 .HasTitle = True 
 .AxisTitle.Caption = "Rebate Amounts" 
 End With 
 End If 
End With
```


## See also


#### Concepts


[Axis Object](axis-object-word.md)

