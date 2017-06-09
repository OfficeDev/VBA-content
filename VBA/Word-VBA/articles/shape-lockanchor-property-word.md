---
title: Shape.LockAnchor Property (Word)
keywords: vbawd10.chm161481006
f1_keywords:
- vbawd10.chm161481006
ms.prod: word
api_name:
- Word.Shape.LockAnchor
ms.assetid: dc153260-5e5d-75f6-c776-481020778cc9
ms.date: 06/08/2017
---


# Shape.LockAnchor Property (Word)

 **True** if the anchor of a **Shape** object is locked to the anchoring range. Read/write **Long** .


## Syntax

 _expression_ . **LockAnchor**

 _expression_ Required. A variable that represents a **[Shape](shape-object-word.md)** object.


## Remarks

When a shape has a locked anchor, you cannot move the shape's anchor by dragging it. The anchor does not move as the shape is moved. A  **Shape** object is anchored to a range of text, but you can position it anywhere on the page. The shape is anchored to the beginning of the first paragraph that contains the anchoring range. A shape will always remain on the same page as its anchor.


## Example

This example creates a new document, adds a shape to it, and then locks the shape's anchor.


```vb
Set myDoc = Documents.Add 
Set myShape = myDoc.Shapes.AddShape(msoShapeBalloon, _ 
 100, 100, 140, 70) 
myShape.LockAnchor = True 
ActiveDocument.ActiveWindow.View.ShowObjectAnchors = True
```

This example returns a message that states the lock status for each shape in the active document.




```
For x = 1 to ActiveDocument.Shapes.Count 
 Msgbox "Shape " &; x &; " is locked - " _ 
 &; ActiveDocument.Shapes(x).LockAnchor 
Next x
```


## See also


#### Concepts


[Shape Object](shape-object-word.md)

