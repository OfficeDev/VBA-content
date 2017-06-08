---
title: PublishObject.RangeEnd Property (PowerPoint)
keywords: vbapp10.chm635006
f1_keywords:
- vbapp10.chm635006
ms.prod: powerpoint
api_name:
- PowerPoint.PublishObject.RangeEnd
ms.assetid: 3edce18e-31c5-4585-9ca5-adb8cbdbca17
ms.date: 06/08/2017
---


# PublishObject.RangeEnd Property (PowerPoint)

Returns or sets the number of the last slide in a range of slides you are publishing as a Web presentation. Read/write.


## Syntax

 _expression_. **RangeEnd**

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

