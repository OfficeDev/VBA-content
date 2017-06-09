---
title: Axis.MajorUnitIsAuto Property (Word)
keywords: vbawd10.chm113049626
f1_keywords:
- vbawd10.chm113049626
ms.prod: word
api_name:
- Word.Axis.MajorUnitIsAuto
ms.assetid: 582059c6-89d4-cd11-e43c-e9c7988fc765
ms.date: 06/08/2017
---


# Axis.MajorUnitIsAuto Property (Word)

 **True** if Microsoft Word calculates the major units for the value axis. Read/write **Boolean** .


## Syntax

 _expression_ . **MajorUnitIsAuto**

 _expression_ A variable that represents an **[Axis](axis-object-word.md)** object.


## Remarks

Setting the  **[MajorUnit](axis-majorunit-property-word.md)** property sets this property to **False** .


## Example

The following example automatically sets the major and minor units for the value axis of the first chart in the active document.


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

