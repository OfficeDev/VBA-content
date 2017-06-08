---
title: Shape.HasSmartArt Property (Word)
keywords: vbawd10.chm161480910
f1_keywords:
- vbawd10.chm161480910
ms.prod: word
api_name:
- Word.Shape.HasSmartArt
ms.assetid: 83aba591-2a42-3ba3-3e73-48ba249b2f7c
ms.date: 06/08/2017
---


# Shape.HasSmartArt Property (Word)

Returns  **True** if there is a SmartArt diagram present on the shape. Read-only.


## Syntax

 _expression_ . **HasSmartArt**

 _expression_ A variable that represents a **[Shape](shape-object-word.md)** object.


## Example

The following code example displays whether or not the first shape in the active document contains SmartArt.


```vb
Dim myShape As Shape 
 
Set myShape = ActiveDocument.Shapes(1) 
 
If myShape.HasSmartArt Then 
 MsgBox "The first shape contains SmartArt." 
Else 
 MsgBox "The first shape contains no SmartArt." 
End If
```


## See also


#### Concepts


[Shape Object](shape-object-word.md)

