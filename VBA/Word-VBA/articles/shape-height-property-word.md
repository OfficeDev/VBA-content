---
title: Shape.Height Property (Word)
keywords: vbawd10.chm161480813
f1_keywords:
- vbawd10.chm161480813
ms.prod: word
api_name:
- Word.Shape.Height
ms.assetid: 3738d3b8-c412-7b42-b430-5082e822eab4
ms.date: 06/08/2017
---


# Shape.Height Property (Word)

Returns or sets the height (in points) of the specified shape. Read/write  **Single** .


## Syntax 

 _expression_ . **Height**

 _expression_ A variable that represents a **[Shape](shape-object-word.md)** object.


## Example

This example inserts a picture as an inline shape and changes the height and width of the image.


```vb
Dim aInLine As Shape 
 
Set aInLine = ActiveDocument.InlineShapes.AddPicture( _ 
 FileName:="C:\Windows\Bubbles.bmp", Range:=Selection.Range) 
 
With aInLine 
 .Height = 100 
 .Width = 200 
End With
```


## See also


#### Concepts


[Shape Object](shape-object-word.md)

