---
title: TextStyle.Ruler Property (PowerPoint)
keywords: vbapp10.chm579003
f1_keywords:
- vbapp10.chm579003
ms.prod: powerpoint
api_name:
- PowerPoint.TextStyle.Ruler
ms.assetid: 01a04a13-d536-72f2-9a7c-07f703e2583c
ms.date: 06/08/2017
---


# TextStyle.Ruler Property (PowerPoint)

Returns a  **[Ruler](ruler-object-powerpoint.md)** object that represents the ruler for the specified text. Read-only.


## Syntax

 _expression_. **Ruler**

 _expression_ A variable that represents a **TextStyle** object.


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


[TextStyle Object](textstyle-object-powerpoint.md)

