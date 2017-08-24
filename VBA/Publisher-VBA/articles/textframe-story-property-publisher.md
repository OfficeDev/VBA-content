---
title: TextFrame.Story Property (Publisher)
keywords: vbapb10.chm3866663
f1_keywords:
- vbapb10.chm3866663
ms.prod: publisher
api_name:
- Publisher.TextFrame.Story
ms.assetid: 7bbe0967-83aa-745b-ad13-8a7dfe61811c
ms.date: 06/08/2017
---


# TextFrame.Story Property (Publisher)

Returns a  **Story** object that represents the story properties in a text range.


## Syntax

 _expression_. **Story**

 _expression_A variable that represents a  **TextFrame** object.


## Example

This example returns the story in the selected text range and, if it is in a text frame, inserts text into the text range.


```vb
Sub AddTextToStory() 
 With Selection.TextRange.Story 
 If .HasTextFrame Then 
 .TextRange.InsertAfter NewText:=vbLf &; "This is a test." 
 End If 
 End With 
End Sub
```


