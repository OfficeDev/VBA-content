---
title: TextFrame.PreviousLinkedTextFrame Property (Publisher)
keywords: vbapb10.chm3866656
f1_keywords:
- vbapb10.chm3866656
ms.prod: publisher
api_name:
- Publisher.TextFrame.PreviousLinkedTextFrame
ms.assetid: 00947ec3-fcff-4451-491b-5b7748ccb74e
ms.date: 06/08/2017
---


# TextFrame.PreviousLinkedTextFrame Property (Publisher)

Returns a  **[TextFrame](textframe-object-publisher.md)** object representing the text frame from which text flows to the specified text frame.


## Syntax

 _expression_. **PreviousLinkedTextFrame**

 _expression_A variable that represents a  **TextFrame** object.


### Return Value

TextFrame


## Remarks

If the specified text frame is not part of a chain of linked frames or is the first in a chain of linked frames, this property returns nothing.


## Example

The following example returns the previously linked text frame of shape three on page one of the active publication and sets its font to Times New Roman.


```vb
Dim txtFrame As TextFrame 
 
Set txtFrame = ActiveDocument.Pages(1) _ 
 .Shapes(3).TextFrame.PreviousLinkedTextFrame 
 
txtFrame.TextRange.Font = "Times New Roman"
```


