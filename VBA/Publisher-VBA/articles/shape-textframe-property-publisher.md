---
title: Shape.TextFrame Property (Publisher)
keywords: vbapb10.chm2228304
f1_keywords:
- vbapb10.chm2228304
ms.prod: publisher
api_name:
- Publisher.Shape.TextFrame
ms.assetid: fc654905-d56b-9a6c-28fa-4b54bf2a8686
ms.date: 06/08/2017
---


# Shape.TextFrame Property (Publisher)

Returns a  **[TextFrame](textframe-object-publisher.md)** object that represents the text in a shape and the properties that control the margins and orientation of the text.


## Syntax

 _expression_. **TextFrame**

 _expression_A variable that represents a  **Shape** object.


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


