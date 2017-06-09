---
title: TextRange.BoundHeight Property (Publisher)
keywords: vbapb10.chm5308436
f1_keywords:
- vbapb10.chm5308436
ms.prod: publisher
api_name:
- Publisher.TextRange.BoundHeight
ms.assetid: 010d3de9-5838-fbf7-fb75-b80a06aafac8
ms.date: 06/08/2017
---


# TextRange.BoundHeight Property (Publisher)

Returns a  **Single** indicating the height, in points, of the bounding box for the specified text range. Read-only.


## Syntax

 _expression_. **BoundHeight**

 _expression_A variable that represents a  **TextRange** object.


### Return Value

Single


## Example

The following example displays the position, width, and height of the bounding box surrounding the text in the first shape on page one of the active publication.


```vb
Dim rngText As TextRange 
Dim strMessage As String 
 
Set rngText = ActiveDocument.Pages(1) _ 
 .Shapes(1).TextFrame.TextRange 
 
With rngText 
 strMessage = "Text frame information" &; vbCrLf _ 
 &; " Distance from left edge of page: " _ 
 &; .BoundLeft &; " points" &; vbCrLf _ 
 &; " Distance from top edge of page: " _ 
 &; .BoundTop &; " points" &; vbCrLf _ 
 &; " Width: " &; .BoundWidth &; " points" &; vbCrLf _ 
 &; " Height: " &; .BoundHeight &; " points" 
End With 
 
MsgBox strMessage
```


