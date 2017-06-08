---
title: TextRange.Copy Method (PowerPoint)
keywords: vbapp10.chm569028
f1_keywords:
- vbapp10.chm569028
ms.prod: powerpoint
api_name:
- PowerPoint.TextRange.Copy
ms.assetid: c8d1edf7-68ef-aaa4-e2db-717263df8dd3
ms.date: 06/08/2017
---


# TextRange.Copy Method (PowerPoint)

Copies the specified object to the Clipboard.


## Syntax

 _expression_. **Copy**

 _expression_ A variable that represents a **TextRange** object.


## Remarks

Use the  **Paste** method to paste the contents of the Clipboard.


## Example

This example copies the text in shape one on slide one in the active presentation to the Clipboard.


```vb
ActivePresentation.Slides(1).Shapes(1).TextFrame.TextRange.Copy
```


## See also


#### Concepts


[TextRange Object](textrange-object-powerpoint.md)

