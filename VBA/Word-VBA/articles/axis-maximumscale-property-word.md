---
title: Axis.MaximumScale Property (Word)
keywords: vbawd10.chm113049628
f1_keywords:
- vbawd10.chm113049628
ms.prod: word
api_name:
- Word.Axis.MaximumScale
ms.assetid: cfd12a67-ef8b-d92c-a9c1-74353754498e
ms.date: 06/08/2017
---


# Axis.MaximumScale Property (Word)

Returns or sets the maximum value on the value axis. Read/write  **Double** .


## Syntax

 _expression_ . **MaximumScale**

 _expression_ A variable that represents an **[Axis](axis-object-word.md)** object.


## Remarks

Setting this property sets the  **[MaximumScaleIsAuto](axis-maximumscaleisauto-property-word.md)** property to **False** .


## Example

The following example sets the minimum and maximum values for the value axis of the first chart in the active document.


```vb
With ActiveDocument.InlineShapes(1) 
 If .HasChart Then 
 With .Chart.Axes(xlValue) 
 .MinimumScale = 10 
 .MaximumScale = 120 
 End With 
 End If 
End With
```


## See also


#### Concepts


[Axis Object](axis-object-word.md)

