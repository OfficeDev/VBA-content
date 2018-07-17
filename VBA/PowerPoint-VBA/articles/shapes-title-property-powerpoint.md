---
title: Shapes.Title Property (PowerPoint)
keywords: vbapp10.chm543020
f1_keywords:
- vbapp10.chm543020
ms.prod: powerpoint
api_name:
- PowerPoint.Shapes.Title
ms.assetid: 61e5f162-d9dd-f8d3-6c15-d5a40c00c10f
ms.date: 06/08/2017
---


# Shapes.Title Property (PowerPoint)

Returns a  **[Shape](shape-object-powerpoint.md)** object that represents the slide title. Read-only.


## Syntax

 _expression_. **Title**

 _expression_ A variable that represents a **Shapes** object.


### Return Value

Shape


## Remarks

You can also use the  **[Item](placeholders-item-method-powerpoint.md)** method of the **[Shapes](shapes-object-powerpoint.md)** or **[Placeholders](placeholders-object-powerpoint.md)** collection to return the slide title.


## Example

This example sets the title text on  `myDocument`.


```vb
Set myDocument = ActivePresentation.Slides(1)

myDocument.Shapes.Title.TextFrame.TextRange.Text = "Welcome!"
```


## See also


#### Concepts


[Shapes Object](shapes-object-powerpoint.md)

