---
title: Axis.Crosses Property (Word)
keywords: vbawd10.chm113049606
f1_keywords:
- vbawd10.chm113049606
ms.prod: word
api_name:
- Word.Axis.Crosses
ms.assetid: 41235c80-55a5-3933-3469-fd95b37ec43c
ms.date: 06/08/2017
---


# Axis.Crosses Property (Word)

Returns or sets the point on the specified axis where the other axis crosses. Read/write  **Long** .


## Syntax

 _expression_ . **Crosses**

 _expression_ A variable that represents an **[Axis](axis-object-word.md)** object.


## Remarks

This property is not available for radar charts. For 3-D charts, this property can be applied only to the value axis and indicates where the plane defined by the category axes crosses the value axis.

You can use this property for both category and value axes. On the category axis,  **xlMinimum** sets the value axis to cross at the first category, and **xlMaximum** sets the value axis to cross at the last category. **xlMinimum** and **xlMaximum** are constants in the **XlAxisCrosses** enumeration.

Note that  **xlMinimum** and **xlMaximum** can have different meanings, depending on the axis.


## Example

The following example sets the value axis in for the first chart in the active document to cross the category axis at the maximum x-axis value.


```vb
With ActiveDocument.InlineShapes(1) 
 If .HasChart Then 
 .Chart.Axes(xlCategory).Crosses = xlMaximum 
 End If 
End With
```


## See also


#### Concepts


[Axis Object](axis-object-word.md)

