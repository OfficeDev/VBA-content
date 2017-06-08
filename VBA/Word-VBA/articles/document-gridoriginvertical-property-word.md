---
title: Document.GridOriginVertical Property (Word)
keywords: vbawd10.chm158007601
f1_keywords:
- vbawd10.chm158007601
ms.prod: word
api_name:
- Word.Document.GridOriginVertical
ms.assetid: 6fd6a060-6f25-b7c6-f4d2-b496c4d2f4b4
ms.date: 06/08/2017
---


# Document.GridOriginVertical Property (Word)

Returns or sets a  **Single** that represents the point, relative to the top of the page, where you want the invisible grid for drawing, moving, and resizing AutoShapes or East Asian characters to begin in the specified document. Read/write.


## Syntax

 _expression_ . **GridOriginVertical**

 _expression_ A variable that represents a **[Document](document-object-word.md)** object.


## Example

This example sets the horizontal and vertical point of origin for the grid, sets the horizontal and vertical distance between gridlines, and then enables the Snap objects to grid feature for the current document.


```vb
With ActiveDocument 
 .GridOriginHorizontal = 80 
 .GridOriginVertical = 90 
 .GridDistanceHorizontal = 9 
 .GridDistanceVertical = 9 
 .SnapToGrid = True 
End With
```


## See also


#### Concepts


[Document Object](document-object-word.md)

