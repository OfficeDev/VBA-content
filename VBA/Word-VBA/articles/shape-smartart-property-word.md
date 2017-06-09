---
title: Shape.SmartArt Property (Word)
keywords: vbawd10.chm161480860
f1_keywords:
- vbawd10.chm161480860
ms.prod: word
api_name:
- Word.Shape.SmartArt
ms.assetid: d2f3fd89-288d-ac1e-18bb-00e2d043d4cd
ms.date: 06/08/2017
---


# Shape.SmartArt Property (Word)

Returns a [SmartArt](http://msdn.microsoft.com/library/24332c9b-87c9-7678-9d9f-9e25f2370afc%28Office.15%29.aspx) object that provides a way to work with the SmartArt associated with the specified shape. Read-only.


## Syntax

 _expression_ . **SmartArt**

 _expression_ A variable that represents a **[Shape](shape-object-word.md)** object.


## Remarks

The  **SmartArt** property provides an entry point for interacting with a SmartArt graphic associated with the shape.


## Example

The following code example adds a SmartArt graphic to the active document.


```vb
Dim myShape As Shape 
Dim mySmartArt As SmartArt 
 
Set myShape = ActiveDocument.Shapes.AddSmartArt(Application.SmartArtLayouts(1), 100, 100, 400, 400) 
Set mySmartArt = myShape.SmartArt
```


## See also


#### Concepts


[Shape Object](shape-object-word.md)

