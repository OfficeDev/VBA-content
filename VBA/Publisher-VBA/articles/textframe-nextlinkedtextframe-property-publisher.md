---
title: TextFrame.NextLinkedTextFrame Property (Publisher)
keywords: vbapb10.chm3866648
f1_keywords:
- vbapb10.chm3866648
ms.prod: publisher
api_name:
- Publisher.TextFrame.NextLinkedTextFrame
ms.assetid: 5ba08ab5-8515-4efe-59a3-79a11f6a7c4e
ms.date: 06/08/2017
---


# TextFrame.NextLinkedTextFrame Property (Publisher)

Returns or sets a  **[TextFrame](textframe-object-publisher.md)** object representing the text frame to which text flows from the specified text frame. Read/write.


## Syntax

 _expression_. **NextLinkedTextFrame**

 _expression_A variable that represents a  **TextFrame** object.


### Return Value

TextFrame


## Remarks

If the specified text frame is not part of a chain of linked frames or is the last in a chain of linked frames, this property returns nothing.


## Example

The following example returns the next linked text frame of shape three on page one of the active publication and sets its font to Times New Roman.


```vb
Dim txtFrame As TextFrame 
 
Set txtFrame = ActiveDocument.Pages(1) _ 
 .Shapes(3).TextFrame.NextLinkedTextFrame 
 
txtFrame.TextRange.Font = "Times New Roman"
```


