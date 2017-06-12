---
title: TextFrame2.Column Property (PowerPoint)
keywords: vbapp10.chm678017
f1_keywords:
- vbapp10.chm678017
ms.prod: powerpoint
api_name:
- PowerPoint.TextFrame2.Column
ms.assetid: d265fd2c-1e96-984d-9b2c-0a792cbf7671
ms.date: 06/08/2017
---


# TextFrame2.Column Property (PowerPoint)

Returns the  **[Column](column-object-powerpoint.md)** object that represents the columns of the specified text frame. Read-only.


## Syntax

 _expression_. **Column**

 _expression_ An expression that returns a **TextFrame2** object.


## Example

The following example shows how to set the number of columns in the text frame of the first shape on slide one to 2.


```vb
Public Sub Column_Example()

    ActivePresentation.Slides(1).Shapes(1).TextFrame2.Column.Number = 2

End Sub
```


## See also


#### Concepts


[TextFrame2 Object](textframe2-object-powerpoint.md)

