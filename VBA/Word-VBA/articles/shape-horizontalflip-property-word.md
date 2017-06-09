---
title: Shape.HorizontalFlip Property (Word)
keywords: vbawd10.chm161480814
f1_keywords:
- vbawd10.chm161480814
ms.prod: word
api_name:
- Word.Shape.HorizontalFlip
ms.assetid: b4bda66d-2826-9f12-1901-d47b824daeda
ms.date: 06/08/2017
---


# Shape.HorizontalFlip Property (Word)

Indicates that a shape has been flipped horizontally. Read-only  **MsoTriState** .


## Syntax

 _expression_ . **HorizontalFlip**

 _expression_ Required. A variable that represents a **[Shape](shape-object-word.md)** object.


## Example

This example restores each shape in the active document to its original state if it has been flipped horizontally or vertically.


```vb
Sub FlipShape() 
 Dim shpFlip As Shape 
 For Each shpFlip In ActiveDocument.Shapes 
 If shpFlip.HorizontalFlip Then shpFlip.Flip msoFlipHorizontal 
 If shpFlip.VerticalFlip Then shpFlip.Flip msoFlipVertical 
 Next 
End Sub
```


## See also


#### Concepts


[Shape Object](shape-object-word.md)

