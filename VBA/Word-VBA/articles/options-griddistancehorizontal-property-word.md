---
title: Options.GridDistanceHorizontal Property (Word)
keywords: vbawd10.chm162988113
f1_keywords:
- vbawd10.chm162988113
ms.prod: word
api_name:
- Word.Options.GridDistanceHorizontal
ms.assetid: 1d28ba4b-ee06-1b1a-e921-2d8d07cab305
ms.date: 06/08/2017
---


# Options.GridDistanceHorizontal Property (Word)

Returns or sets the amount of horizontal space between the invisible gridlines that Word uses when you draw, move, and resize AutoShapes or East Asian characters in new documents. Read/write  **Single** .


## Syntax

 _expression_ . **GridDistanceHorizontal**

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

