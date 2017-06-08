---
title: DropCap.Span Property (Publisher)
keywords: vbapb10.chm5505033
f1_keywords:
- vbapb10.chm5505033
ms.prod: publisher
api_name:
- Publisher.DropCap.Span
ms.assetid: 00c51e48-5bbc-13e9-2d0c-e8993f753bbe
ms.date: 06/08/2017
---


# DropCap.Span Property (Publisher)

Returns or sets a  **Long** that represents the number of letters included in the specified dropped capital letter. Read/write.


## Syntax

 _expression_. **Span**

 _expression_A variable that represents a  **DropCap** object.


### Return Value

Long


## Example

This example creates a custom dropped capital letter that is five lines high, spans the first three characters of the paragraphs in the text range, and is raised one line above the first line.


```vb
Sub SetDropCapSpan() 
 With ActiveDocument.Pages(1).Shapes(1) _ 
 .TextFrame.TextRange.DropCap 
 .Size = 5 
 .Span = 3 
 .LinesUp = 1 
 End With 
End Sub
```


