---
title: Axis.HasMinorGridlines Property (Word)
keywords: vbawd10.chm113049613
f1_keywords:
- vbawd10.chm113049613
ms.prod: word
api_name:
- Word.Axis.HasMinorGridlines
ms.assetid: f835dab5-1256-bd4c-0219-2e3016120d18
ms.date: 06/08/2017
---


# Axis.HasMinorGridlines Property (Word)

 **True** if the axis has minor gridlines. Read/write **Boolean** .


## Syntax

 _expression_ . **HasMinorGridlines**

 _expression_ A variable that represents an **[Axis](axis-object-word.md)** object.


## Remarks

Only axes in the primary axis group can have gridlines.


## Example

The following example sets the color of the minor gridlines for the value axis of the first chart in the active document.


```vb
With ActiveDocument.InlineShapes(1) 
 If .HasChart Then 
 With .Chart.Axes(xlValue) 
 If .HasMinorGridlines Then 
 ' Set the color to green. 
 .MinorGridlines.Border.ColorIndex = 4 
 End If 
 End With 
 End If 
End With 

```


## See also


#### Concepts


[Axis Object](axis-object-word.md)

