---
title: TextFrame.Ruler Property (PowerPoint)
keywords: vbapp10.chm558009
f1_keywords:
- vbapp10.chm558009
ms.prod: powerpoint
api_name:
- PowerPoint.TextFrame.Ruler
ms.assetid: 496ef8d2-b8c5-71a6-93d4-23e0a8d171f3
ms.date: 06/08/2017
---


# TextFrame.Ruler Property (PowerPoint)

Returns a  **[Ruler](ruler-object-powerpoint.md)** object that represents the ruler for the specified text. Read-only.


## Syntax

 _expression_. **Ruler**

 _expression_ A variable that represents a **TextFrame** object.


### Return Value

Ruler


## Example

This example sets a left-aligned tab stop at 2 inches (144 points) for the text in shape two on  `myDocument`.


```vb
Set myDocument = ActivePresentation.Slides(1)

myDocument.Shapes(2).TextFrame.Ruler.TabStops _
    .Add ppTabStopLeft, 144
```


## See also


#### Concepts


[TextFrame Object](textframe-object-powerpoint.md)

