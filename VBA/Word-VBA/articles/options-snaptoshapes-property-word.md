---
title: Options.SnapToShapes Property (Word)
keywords: vbawd10.chm162988112
f1_keywords:
- vbawd10.chm162988112
ms.prod: word
api_name:
- Word.Options.SnapToShapes
ms.assetid: 7433f9ec-d67b-eaaf-7ae5-129bf7aba7ff
ms.date: 06/08/2017
---


# Options.SnapToShapes Property (Word)

 **True** if Word automatically aligns AutoShapes or East Asian characters with invisible gridlines that go through the vertical and horizontal edges of other AutoShapes or East Asian characters. Read/write **Boolean** .


## Syntax

 _expression_ . **SnapToShapes**

 _expression_ A variable that represents an **[Options](options-object-word.md)** collection.


## Remarks

This property creates additional invisible gridlines for each AutoShape.  **SnapToShapes** works independently of the **SnapToGrid** property.


## Example

This example sets Word to automatically align AutoShapes with invisible gridlines that go through the vertical and horizontal edges of other AutoShapes in a new document.


```vb
Options.SnapToShapes = True 
Documents.Add
```


## See also


#### Concepts


[Options Object](options-object-word.md)

