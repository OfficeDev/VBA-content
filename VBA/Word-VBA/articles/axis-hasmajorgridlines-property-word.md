---
title: Axis.HasMajorGridlines Property (Word)
keywords: vbawd10.chm113049611
f1_keywords:
- vbawd10.chm113049611
ms.prod: word
api_name:
- Word.Axis.HasMajorGridlines
ms.assetid: bd207374-f9b1-ed1d-f309-30c07ebf1e70
ms.date: 06/08/2017
---


# Axis.HasMajorGridlines Property (Word)

 **True** if the axis has major gridlines. Read/write **Boolean** .


## Syntax

 _expression_ . **HasMajorGridlines**

 _expression_ A variable that represents an **[Axis](axis-object-word.md)** object.


## Remarks

Only axes in the primary axis group can have gridlines.


## Example

The following example sets the color of the major gridlines for the value axis of the first chart in the active document.


```vb
With ActiveDocument.InlineShapes(1) 
 If .HasChart Then 
 With .Chart.Axes(xlValue) 
 If .HasMajorGridlines Then 
 ' Set the color to red. 
 .MajorGridlines.Border.ColorIndex = 3 
 End If 
 End With 
 End If 
End With 

```


## See also


#### Concepts


[Axis Object](axis-object-word.md)

