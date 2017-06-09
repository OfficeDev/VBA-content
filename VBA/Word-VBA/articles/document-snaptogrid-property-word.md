---
title: Document.SnapToGrid Property (Word)
keywords: vbawd10.chm158007596
f1_keywords:
- vbawd10.chm158007596
ms.prod: word
api_name:
- Word.Document.SnapToGrid
ms.assetid: 7aa03a0d-65f2-725b-37fe-8a421fb1e9f7
ms.date: 06/08/2017
---


# Document.SnapToGrid Property (Word)

 **True** if AutoShapes or East Asian characters are automatically aligned with an invisible grid when they are drawn, moved, or resized in the specified document. Read/write **Boolean** .


## Syntax

 _expression_ . **SnapToGrid**

 _expression_ A variable that represents a **[Document](document-object-word.md)** object.


## Remarks

You can temporarily override this setting by pressing ALT while drawing, moving, or resizing an AutoShape.


## Example

This example sets Microsoft Word to automatically align East Asian characters with the invisible grid in the current document.


```vb
ActiveDocument.SnapToGrid = True
```


## See also


#### Concepts


[Document Object](document-object-word.md)

