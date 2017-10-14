---
title: Point.ApplyPictToEnd Property (Word)
keywords: vbawd10.chm262145661
f1_keywords:
- vbawd10.chm262145661
ms.prod: word
api_name:
- Word.Point.ApplyPictToEnd
ms.assetid: 4755d10d-5844-0274-d0e5-fc90e7c2e779
ms.date: 06/08/2017
---


# Point.ApplyPictToEnd Property (Word)

 **True** if a picture is applied to the end of the point or all points in the series. Read/write **Boolean** .


## Syntax

 _expression_ . **ApplyPictToEnd**

 _expression_ A variable that represents a **[Point](point-object-word.md)** object.


## Example

The following example applies pictures to the end of all points in the first series of the first chart in the active document. The series must already have pictures applied to it (setting this property changes the picture orientation).


```vb
With ActiveDocument.InlineShapes(1) 
 If .HasChart Then 
 .Chart.SeriesCollection(1).ApplyPictToEnd = True 
 End If 
End With
```


## See also


#### Concepts


[Point Object](point-object-word.md)

