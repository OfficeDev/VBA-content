---
title: TabStops.DefaultSpacing Property (PowerPoint)
keywords: vbapp10.chm573004
f1_keywords:
- vbapp10.chm573004
ms.prod: powerpoint
api_name:
- PowerPoint.TabStops.DefaultSpacing
ms.assetid: f404f50d-ae85-3310-a478-6800d39fb582
ms.date: 06/08/2017
---


# TabStops.DefaultSpacing Property (PowerPoint)

Returns or sets the default tab-stop spacing for the specified text, in points. Read/write.


## Syntax

 _expression_. **DefaultSpacing**

 _expression_ A variable that represents a **TabStops** object.


### Return Value

Single


## Example

This example sets the default tab-stop spacing to 0.5 inch (36 points) for the text in shape two on slide one in the active presentation.


```vb
Application.ActivePresentation.Slides(1).Shapes(2).TextFrame _
    .Ruler.TabStops.DefaultSpacing = 36
```


## See also


#### Concepts


[TabStops Object](tabstops-object-powerpoint.md)

