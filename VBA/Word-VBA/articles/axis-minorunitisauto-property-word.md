---
title: Axis.MinorUnitIsAuto Property (Word)
keywords: vbawd10.chm113049641
f1_keywords:
- vbawd10.chm113049641
ms.prod: word
api_name:
- Word.Axis.MinorUnitIsAuto
ms.assetid: 6ea041c2-b1f3-73b6-f9b4-707edc611ba4
ms.date: 06/08/2017
---


# Axis.MinorUnitIsAuto Property (Word)

 **True** if Microsoft Word calculates minor units for the value axis. Read/write **Boolean** .


## Syntax

 _expression_ . **MinorUnitIsAuto**

 _expression_ A variable that represents an **[Axis](axis-object-word.md)** object.


## Remarks

Setting the  **[MinorUnit](axis-minorunit-property-word.md)** property sets this property to **False** .


## Example

The following example automatically calculates major and minor units for the value axis of the first chart in the active document.


```vb
With ActiveDocument.InlineShapes(1) 
 If .HasChart Then 
 With .Chart.Axes(xlValue) 
 .MajorUnitIsAuto = True 
 .MinorUnitIsAuto = True 
 End With 
 End If 
End With
```


## See also


#### Concepts


[Axis Object](axis-object-word.md)

