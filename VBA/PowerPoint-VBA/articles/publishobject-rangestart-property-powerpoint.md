---
title: PublishObject.RangeStart Property (PowerPoint)
keywords: vbapp10.chm635005
f1_keywords:
- vbapp10.chm635005
ms.prod: powerpoint
api_name:
- PowerPoint.PublishObject.RangeStart
ms.assetid: c7b576f4-f001-994a-ef36-0ed9402960a2
ms.date: 06/08/2017
---


# PublishObject.RangeStart Property (PowerPoint)

Returns or sets the number of the first slide in a range of slides you are publishing as a Web presentation. Read/write.


## Syntax

 _expression_. **RangeStart**

 _expression_ A variable that represents a **PublishObject** object.


### Return Value

Integer


## Example

This example publishes slides three through five of the active presentation to HTML. It names the published presentation Mallard.htm.


```vb
With ActivePresentation.PublishObjects(1)

    .FileName = "C:\Test\Mallard.htm"

    .SourceType = ppPublishSlideRange

    .RangeStart = 3

    .RangeEnd = 5

    .Publish

End With
```


## See also


#### Concepts


[PublishObject Object](publishobject-object-powerpoint.md)

