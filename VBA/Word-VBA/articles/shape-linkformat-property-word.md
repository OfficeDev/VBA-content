---
title: Shape.LinkFormat Property (Word)
keywords: vbawd10.chm161481206
f1_keywords:
- vbawd10.chm161481206
ms.prod: word
api_name:
- Word.Shape.LinkFormat
ms.assetid: 68645111-3036-da95-eab7-3d78a6896e51
ms.date: 06/08/2017
---


# Shape.LinkFormat Property (Word)

Returns a  **LinkFormat** object that represents the link options of a shape that is linked to a file. Read/only.


## Syntax

 _expression_ . **LinkFormat**

 _expression_ A variable that represents a **[Shape](shape-object-word.md)** object.


## Example

This example inserts a graphic as an inline shape (using an INCLUDEPICTURE field) and then displays the source name (Tiles.bmp).


```vb
Set iShape = ActiveDocument.InlineShapes _ 
 .AddPicture(FileName:="C:\windows\Tiles.bmp", _ 
 LinkToFile:=True, SaveWithDocument:=False, _ 
 Range:=Selection.Range) 
MsgBox iShape.LinkFormat.SourceName
```


## See also


#### Concepts


[Shape Object](shape-object-word.md)

