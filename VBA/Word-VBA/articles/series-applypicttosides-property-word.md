---
title: Series.ApplyPictToSides Property (Word)
keywords: vbawd10.chm123733627
f1_keywords:
- vbawd10.chm123733627
ms.prod: word
api_name:
- Word.Series.ApplyPictToSides
ms.assetid: b8277abd-64c6-2b1c-23e6-5ff8c21619fc
ms.date: 06/08/2017
---


# Series.ApplyPictToSides Property (Word)

 **True** if a picture is applied to the sides of the point or all points in the series. Read/write **Boolean** .


## Syntax

 _expression_ . **ApplyPictToSides**

 _expression_ A variable that represents a **[Series](series-object-word.md)** object.


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


[Series Object](series-object-word.md)

