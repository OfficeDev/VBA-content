---
title: TextRange.DropCap Property (Publisher)
keywords: vbapb10.chm5308472
f1_keywords:
- vbapb10.chm5308472
ms.prod: publisher
api_name:
- Publisher.TextRange.DropCap
ms.assetid: a5c29dd4-62f4-39fb-4b76-390d62bd8e32
ms.date: 06/08/2017
---


# TextRange.DropCap Property (Publisher)

Returns a  **[DropCap](dropcap-object-publisher.md)** object that represents a dropped capital letter for the paragraphs in the specified text frame.


## Syntax

 _expression_. **DropCap**

 _expression_A variable that represents a  **TextRange** object.


### Return Value

DropCap


## Example

This example applies a custom dropped capital that is three lines high and spans the first three characters of each paragraph in the specified text frame.


```vb
Sub SetDropCap() 
 With ActiveDocument.Pages(1).Shapes(1).TextFrame.TextRange 
 .DropCap.ApplyCustomDropCap FontName:="Snap ITC", _ 
 Bold:=True, Size:=3, Span:=3 
 With .ParagraphFormat 
 .SpaceBefore = 6 
 .SpaceAfter = 6 
 End With 
 End With 
End Sub
```


