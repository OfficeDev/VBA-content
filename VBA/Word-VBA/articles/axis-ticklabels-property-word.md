---
title: Axis.TickLabels Property (Word)
keywords: vbawd10.chm113049650
f1_keywords:
- vbawd10.chm113049650
ms.prod: word
api_name:
- Word.Axis.TickLabels
ms.assetid: 5c363e25-71e3-4f89-bcd3-612855000f53
ms.date: 06/08/2017
---


# Axis.TickLabels Property (Word)

Returns the tick-mark labels for the specified axis. Read-only  **[TickLabels](ticklabels-object-word.md)** .


## Syntax

 _expression_ . **TickLabels**

 _expression_ A variable that represents an **[Axis](axis-object-word.md)** object.


## Example

The following example sets the color of the tick-mark label font for the value axis of the first chart in the active document.


```vb
With ActiveDocument.InlineShapes(1) 
 If .HasChart Then 
 .Chart.Axes(xlValue).TickLabels.Font.ColorIndex = 3 
 End If 
End With
```


## See also


#### Concepts


[Axis Object](axis-object-word.md)

