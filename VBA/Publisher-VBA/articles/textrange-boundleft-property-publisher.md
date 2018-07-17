---
title: TextRange.BoundLeft Property (Publisher)
keywords: vbapb10.chm5308435
f1_keywords:
- vbapb10.chm5308435
ms.prod: publisher
api_name:
- Publisher.TextRange.BoundLeft
ms.assetid: 1ad36906-3dbf-9158-173b-b9047910f6d2
ms.date: 06/08/2017
---


# TextRange.BoundLeft Property (Publisher)

Returns a  **Single** indicating the distance, in points, from the left edge of the leftmost page to the left edge of the bounding box for the specified text range. Read-only.


## Syntax

 _expression_. **BoundLeft**

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


