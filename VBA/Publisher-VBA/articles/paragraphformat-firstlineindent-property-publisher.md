---
title: ParagraphFormat.FirstLineIndent Property (Publisher)
keywords: vbapb10.chm5439493
f1_keywords:
- vbapb10.chm5439493
ms.prod: publisher
api_name:
- Publisher.ParagraphFormat.FirstLineIndent
ms.assetid: 4966b30e-7629-b66d-0870-ada91c3af4f3
ms.date: 06/08/2017
---


# ParagraphFormat.FirstLineIndent Property (Publisher)

Returns or sets a  **Variant** that represents the amount of space (measured in points) to indent the first line in a paragraph. Read/write.


## Syntax

 _expression_. **FirstLineIndent**

 _expression_A variable that represents a  **ParagraphFormat** object.


### Return Value

Variant


## Example

This example creates a text box, fills it with text, and indents the first line of every paragraph a half inch.


```vb
Sub IndentFirstLines() 
 Dim intCount As Integer 
 With ActiveDocument.Pages(1).Shapes _ 
 .AddTextbox(Orientation:=pbTextOrientationHorizontal, _ 
 Left:=100, Top:=100, Width:=100, Height:=100) _ 
 .TextFrame.TextRange 
 For intCount = 1 To 10 
 .InsertAfter NewText:="This is a test. " 
 Next intCount 
 .ParagraphFormat.FirstLineIndent = InchesToPoints(0.5) 
 End With 
End Sub
```


