---
title: Series.ApplyPictToEnd Property (Word)
keywords: vbawd10.chm123733629
f1_keywords:
- vbawd10.chm123733629
ms.prod: word
api_name:
- Word.Series.ApplyPictToEnd
ms.assetid: d21d40d6-7d66-7513-a225-e151e64c4188
ms.date: 06/08/2017
---


# Series.ApplyPictToEnd Property (Word)

 **True** if a picture is applied to the end of the point or all points in the series. Read/write **Boolean** .


## Syntax

 _expression_ . **ApplyPictToEnd**

 _expression_ A variable that represents a **[Series](series-object-word.md)** object.


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


[Series Object](series-object-word.md)

