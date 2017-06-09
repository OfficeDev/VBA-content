---
title: Axis.BaseUnitIsAuto Property (Word)
keywords: vbawd10.chm113049659
f1_keywords:
- vbawd10.chm113049659
ms.prod: word
api_name:
- Word.Axis.BaseUnitIsAuto
ms.assetid: 7dcfd41c-c35d-5a61-55bd-e7e675fb589c
ms.date: 06/08/2017
---


# Axis.BaseUnitIsAuto Property (Word)

 **True** if Microsoft Word chooses appropriate base units for the specified category axis. The default is **True** . Read/write **Boolean** .


## Syntax

 _expression_ . **BaseUnitIsAuto**

 _expression_ A variable that represents an **[Axis](axis-object-word.md)** object.


## Remarks

You cannot set this property for a value axis.


## Example

The following example sets the category axis for the first chart in the active document to use a time scale, with the base unit automatically chosen by Word.


```vb
 
With ActiveDocument.InlineShapes(1) 
 If .HasChart Then 
 With .Chart 
 .Axes(xlCategory).CategoryType = xlTimeScale 
 .Axes(xlCategory).BaseUnitIsAuto = True 
 End With 
 End If 
End With
```


## See also


#### Concepts


[Axis Object](axis-object-word.md)

