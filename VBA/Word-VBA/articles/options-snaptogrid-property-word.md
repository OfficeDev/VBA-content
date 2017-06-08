---
title: Options.SnapToGrid Property (Word)
keywords: vbawd10.chm162988111
f1_keywords:
- vbawd10.chm162988111
ms.prod: word
api_name:
- Word.Options.SnapToGrid
ms.assetid: 253c0e7a-02d3-30da-ebe6-60f73894a421
ms.date: 06/08/2017
---


# Options.SnapToGrid Property (Word)

 **True** if AutoShapes or East Asian characters are automatically aligned with an invisible grid when they are drawn, moved, or resized. Read/write **Boolean** .


## Syntax

 _expression_ . **SnapToGrid**

 _expression_ A variable that represents an **[Options](options-object-word.md)** collection.


## Remarks

You can temporarily override this setting by pressing ALT while drawing, moving, or resizing an AutoShape.


## Example

This example sets Word so that AutoShapes are automatically aligned with the invisible grid in a new document.


```vb
Options.SnapToGrid = True 
Documents.Add
```

This example returns the status of the  **Snap to grid** option in the **Snap to Grid** dialog box.




```
Temp = Options.SnapToGrid
```


## See also


#### Concepts


[Options Object](options-object-word.md)

