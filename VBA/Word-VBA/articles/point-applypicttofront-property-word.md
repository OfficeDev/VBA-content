---
title: Point.ApplyPictToFront Property (Word)
keywords: vbawd10.chm262145660
f1_keywords:
- vbawd10.chm262145660
ms.prod: word
api_name:
- Word.Point.ApplyPictToFront
ms.assetid: 19a8549e-0d5d-7537-a32b-be1e1ae7178e
ms.date: 06/08/2017
---


# Point.ApplyPictToFront Property (Word)

 **True** if a picture is applied to the front of the point or all points in the series. Read/write **Boolean** .


## Syntax

 _expression_ . **ApplyPictToFront**

 _expression_ A variable that represents a **[Point](point-object-word.md)** object.


## Example

The following example applies pictures to the front of all points in the first series of the first chart in the active document. The series must already have pictures applied to it (setting this property changes the picture orientation).


```vb
With ActiveDocument.InlineShapes(1) 
 If .HasChart Then 
 .Chart.SeriesCollection(1).ApplyPictToFront = True 
 End If 
End With
```


## See also


#### Concepts


[Point Object](point-object-word.md)

