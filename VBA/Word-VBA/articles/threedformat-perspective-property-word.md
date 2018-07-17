---
title: ThreeDFormat.Perspective Property (Word)
keywords: vbawd10.chm164626535
f1_keywords:
- vbawd10.chm164626535
ms.prod: word
api_name:
- Word.ThreeDFormat.Perspective
ms.assetid: 89d627c6-43d8-35d3-ad01-e6fc7f3e5142
ms.date: 06/08/2017
---


# ThreeDFormat.Perspective Property (Word)

 **MsoTrue** if the extrusion appears in perspective — that is, if the walls of the extrusion narrow toward a vanishing point. **MsoFalse** if the extrusion is a parallel, or orthographic, projection — that is, if the walls don't narrow toward a vanishing point. Read/write **MsoTriState** .


## Syntax

 _expression_ . **Perspective**

 _expression_ Required. A variable that represents a **[ThreeDFormat](threedformat-object-word.md)** object.


## Example

This example sets the extrusion depth for shape one on myDocument to 100 points and specifies that the extrusion be parallel, or orthographic.


```vb
Set myDocument = ActiveDocument 
With myDocument.Shapes(1).ThreeD 
 .Visible = True 
 .Depth = 100 
 .Perspective = msoFalse 
End With
```


## See also


#### Concepts


[ThreeDFormat Object](threedformat-object-word.md)

