---
title: Series.ApplyPictToFront Property (Word)
keywords: vbawd10.chm123733628
f1_keywords:
- vbawd10.chm123733628
ms.prod: word
api_name:
- Word.Series.ApplyPictToFront
ms.assetid: 85390115-161c-bc63-fbfb-25d793aff79d
ms.date: 06/08/2017
---


# Series.ApplyPictToFront Property (Word)

 **True** if a picture is applied to the front of the point or all points in the series. Read/write **Boolean** .


## Syntax

 _expression_ . **ApplyPictToFront**

 _expression_ A variable that represents a **[Series](series-object-word.md)** object.


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


[Series Object](series-object-word.md)

