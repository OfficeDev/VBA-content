---
title: ShapeRange.TextFrame Property (Publisher)
keywords: vbapb10.chm2293840
f1_keywords:
- vbapb10.chm2293840
ms.prod: publisher
api_name:
- Publisher.ShapeRange.TextFrame
ms.assetid: 2dbb7fb4-3ae4-d4c1-8b7e-3e087e32a96f
ms.date: 06/08/2017
---


# ShapeRange.TextFrame Property (Publisher)

Returns a  **[TextFrame](textframe-object-publisher.md)** object that represents the text in a shape and the properties that control the margins and orientation of the text.


## Syntax

 _expression_. **TextFrame**

 _expression_A variable that represents a  **ShapeRange** object.


## Example

The following example adds text to the text frame of shape one in the active publication, and then formats the new text. This example assumes there is at least one shape on the first page of the active publication.


```vb
Sub AddTextToTextFrame() 
 With ActiveDocument.Pages(1).Shapes(1).TextFrame.TextRange 
 .Text = "My Text" 
 With .Font 
 .Bold = msoTrue 
 .Size = 25 
 .Name = "Arial" 
 End With 
 End With 
End Sub
```


