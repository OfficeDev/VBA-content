---
title: Options.GridDistanceVertical Property (Word)
keywords: vbawd10.chm162988114
f1_keywords:
- vbawd10.chm162988114
ms.prod: word
api_name:
- Word.Options.GridDistanceVertical
ms.assetid: 6da3a2b5-3961-f8ba-c61f-ec1d1d2ea971
ms.date: 06/08/2017
---


# Options.GridDistanceVertical Property (Word)

Returns or sets the amount of vertical space between the invisible gridlines that Word uses when you draw, move, and resize AutoShapes or East Asian characters in new documents. Read/write  **Single** .


## Syntax

 _expression_ . **GridDistanceVertical**

 _expression_ A variable that represents an **[Options](options-object-word.md)** collection.


## Example

This example sets the horizontal and vertical distance between gridlines and then enables the Snap objects to grid feature for a new document.


```vb
With Options 
 .GridDistanceHorizontal = InchesToPoints(0.2) 
 .GridDistanceVertical = InchesToPoints(0.2) 
 .SnapToGrid = True 
End With 
Documents.Add
```


## See also


#### Concepts


[Options Object](options-object-word.md)

