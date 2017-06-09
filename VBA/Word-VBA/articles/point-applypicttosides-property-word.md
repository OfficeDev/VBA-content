---
title: Point.ApplyPictToSides Property (Word)
keywords: vbawd10.chm262145659
f1_keywords:
- vbawd10.chm262145659
ms.prod: word
api_name:
- Word.Point.ApplyPictToSides
ms.assetid: 6f12c8f9-ec8f-18ca-9e77-ddc09a9be167
ms.date: 06/08/2017
---


# Point.ApplyPictToSides Property (Word)

 **True** if a picture is applied to the sides of the point or all points in the series. Read/write **Boolean** .


## Syntax

 _expression_ . **ApplyPictToSides**

 _expression_ A variable that represents a **[Point](point-object-word.md)** object.


## Example

The following example applies pictures to the sides of all points in the first series of the first chart in the active document. The series must already have pictures applied to it (setting this property changes the picture orientation).


```vb
With ActiveDocument.InlineShapes(1) 
 If .HasChart Then 
 .Chart.SeriesCollection(1).ApplyPictToSides = True 
 End If 
End With
```


## See also


#### Concepts


[Point Object](point-object-word.md)

