---
title: Axis.AxisBetweenCategories Property (Word)
keywords: vbawd10.chm113049600
f1_keywords:
- vbawd10.chm113049600
ms.prod: word
api_name:
- Word.Axis.AxisBetweenCategories
ms.assetid: b99e83a2-5540-e69d-402c-224612f8e568
ms.date: 06/08/2017
---


# Axis.AxisBetweenCategories Property (Word)

 **True** if the value axis crosses the category axis between categories. Read/write **Boolean** .


## Syntax

 _expression_ . **AxisBetweenCategories**

 _expression_ A variable that represents an **[Axis](axis-object-word.md)** object.


## Remarks

This property applies only to category axes, and it does not apply to 3-D charts.


## Example

The following example causes the value axis for the first chart in the active document to cross the category axis between categories.


```vb
With ActiveDocument.InlineShapes(1) 
 If .HasChart Then 
 .Chart.Axes(xlCategory). _ 
 AxisBetweenCategories = True 
 End If 
End With
```


## See also


#### Concepts


[Axis Object](axis-object-word.md)

