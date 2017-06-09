---
title: Font.Size Property (PowerPoint)
keywords: vbapp10.chm575014
f1_keywords:
- vbapp10.chm575014
ms.prod: powerpoint
api_name:
- PowerPoint.Font.Size
ms.assetid: dd56a4e9-20c7-b38d-0d0e-82e5326d51c4
ms.date: 06/08/2017
---


# Font.Size Property (PowerPoint)

Returns or sets the character size, in points. Read/write.


## Syntax

 _expression_. **Size**

 _expression_ A variable that represents a **Font** object.


### Return Value

Single


## Example

This example sets the size of the text attached to shape one on slide one to 24 points.


```vb
Application.ActivePresentation.Slides(1) _
    .Shapes(1).TextFrame.TextRange.Font _
    .Size = 24
```


## See also


#### Concepts


[Font Object](font-object-powerpoint.md)

