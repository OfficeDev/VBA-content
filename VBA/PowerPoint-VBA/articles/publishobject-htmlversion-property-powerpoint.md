---
title: PublishObject.HTMLVersion Property (PowerPoint)
keywords: vbapp10.chm635003
f1_keywords:
- vbapp10.chm635003
ms.prod: powerpoint
api_name:
- PowerPoint.PublishObject.HTMLVersion
ms.assetid: 39d6328d-f361-d2ae-34fd-03543c9883a7
ms.date: 06/08/2017
---


# PublishObject.HTMLVersion Property (PowerPoint)

Returns or sets the version of HTML for a published presentation. Read/write.


## Syntax

 _expression_. **HTMLVersion**

 _expression_ A variable that represents a **PublishObject** object.


### Return Value

PpHTMLVersion


## Remarks

The value returned by the  **HTMLVersion** property can be one of these **PpHTMLVersion** constants. The default is **ppHTMLv4**.


||
|:-----|
|**ppHTMLAutodetect**|
|**ppHTMLDual**|
|**ppHTMLv3**|
|**ppHTMLv4**|

## Example

This example publishes slides three through five of the active presentation in HTML version 3.0. It names the published presentation Mallard.htm.


```vb
With ActivePresentation.PublishObjects(1)

    .FileName = "C:\Test\Mallard.htm"

    .SourceType = ppPublishSlideRange

    .RangeStart = 3

    .RangeEnd = 5

    .HTMLVersion = ppHTMLv3

    .Publish

End With
```


## See also


#### Concepts


[PublishObject Object](publishobject-object-powerpoint.md)

