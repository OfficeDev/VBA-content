---
title: DropCap.Size Property (Publisher)
keywords: vbapb10.chm5505032
f1_keywords:
- vbapb10.chm5505032
ms.prod: publisher
api_name:
- Publisher.DropCap.Size
ms.assetid: c8111c4f-7b70-76ba-5c8e-acaeb4c90be7
ms.date: 06/08/2017
---


# DropCap.Size Property (Publisher)

Returns or sets a  **Long** that represents the number of lines high to format a dropped capital letter. Read/write.


## Syntax

 _expression_. **Size**

 _expression_A variable that represents a  **DropCap** object.


## Example

This example formats a drop cap for the specified text range that is five lines high.


```vb
Sub RaisedDropCap() 
 Dim intCount As Integer 
 With ActiveDocument.Pages(1).Shapes _ 
 .AddTextbox(Orientation:=pbTextOrientationHorizontal, _ 
 Left:=100, Top:=100, Width:=100, Height:=100) _ 
 .TextFrame.TextRange 
 For intCount = 1 To 10 
 .InsertAfter NewText:="This is a test. " 
 Next intCount 
 With .DropCap 
 .Size = 5 
 .LinesUp = 2 
 End With 
 End With 
End Sub
```


