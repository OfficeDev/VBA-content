---
title: Axis.MinimumScale Property (Word)
keywords: vbawd10.chm113049632
f1_keywords:
- vbawd10.chm113049632
ms.prod: word
api_name:
- Word.Axis.MinimumScale
ms.assetid: ccc3eb87-4839-5952-263b-00aad68b3521
ms.date: 06/08/2017
---


# Axis.MinimumScale Property (Word)

Returns or sets the minimum value on the value axis. Read/write  **Double** .


## Syntax

 _expression_ . **MinimumScale**

 _expression_ A variable that represents an **[Axis](axis-object-word.md)** object.


## Remarks

Setting this property sets the  **[MinimumScaleIsAuto](axis-minimumscaleisauto-property-word.md)** property to **False** .


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

