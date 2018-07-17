---
title: Point.HasDataLabel Property (Word)
keywords: vbawd10.chm262144077
f1_keywords:
- vbawd10.chm262144077
ms.prod: word
api_name:
- Word.Point.HasDataLabel
ms.assetid: 0b386560-702f-b9d6-b8a0-b5d67189d979
ms.date: 06/08/2017
---


# Point.HasDataLabel Property (Word)

 **True** if the point has a data label. Read/write **Boolean** .


## Syntax

 _expression_ . **HasDataLabel**

 _expression_ A variable that represents a **[Point](point-object-word.md)** object.


## Example

The following example enables the data label for point seven in series three for the first chart in the active document, and then it sets the data label color to blue.


```vb
With ActiveDocument.InlineShapes(1) 
 If .HasChart Then 
 With Chart.SeriesCollection(3).Points(7) 
 .HasDataLabel = True 
 .ApplyDataLabels Type:=xlValue 
 .DataLabel.Font.ColorIndex = 5 
 End With 
 End If 
End With
```


## See also


#### Concepts


[Point Object](point-object-word.md)

