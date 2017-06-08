---
title: Story.TextFrame Property (Publisher)
keywords: vbapb10.chm5832709
f1_keywords:
- vbapb10.chm5832709
ms.prod: publisher
api_name:
- Publisher.Story.TextFrame
ms.assetid: bb6ce510-068c-27c2-9df0-a709ab46db2e
ms.date: 06/08/2017
---


# Story.TextFrame Property (Publisher)

Returns a  **[TextFrame](textframe-object-publisher.md)** object that represents the text in a shape and the properties that control the margins and orientation of the text.


## Syntax

 _expression_. **TextFrame**

 _expression_A variable that represents a  **Story** object.


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


