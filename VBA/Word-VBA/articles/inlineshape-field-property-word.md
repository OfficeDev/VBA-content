---
title: InlineShape.Field Property (Word)
keywords: vbawd10.chm162004996
f1_keywords:
- vbawd10.chm162004996
ms.prod: word
api_name:
- Word.InlineShape.Field
ms.assetid: cc74cfc7-5182-1447-208b-1e6166ffc915
ms.date: 06/08/2017
---


# InlineShape.Field Property (Word)

Returns a  **Field** object that represents the field associated with the specified inline shape. Read-only.


## Syntax

 _expression_ . **Field**

 _expression_ A variable that represents an **[InlineShape](inlineshape-object-word.md)** object.


## Remarks

Use the  **Fields** property to return the **Fields** collection.


## Example

This example inserts a graphic as an inline shape (using an INCLUDEPICTURE field) and then displays the shape's field code.


```vb
Dim iShapeNew As InlineShape 
 
Set iShapeNew = _ 
 ActiveDocument.InlineShapes _ 
 .AddPicture(FileName:="C:\Windows\Tiles.bmp", _ 
 LinkToFile:=True, SaveWithDocument:=False, _ 
 Range:=Selection.Range) 
 
MsgBox iShapeNew.Field.Code.Text
```


## See also


#### Concepts


[InlineShape Object](inlineshape-object-word.md)

