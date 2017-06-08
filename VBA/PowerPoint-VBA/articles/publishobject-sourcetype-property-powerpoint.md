---
title: PublishObject.SourceType Property (PowerPoint)
keywords: vbapp10.chm635004
f1_keywords:
- vbapp10.chm635004
ms.prod: powerpoint
api_name:
- PowerPoint.PublishObject.SourceType
ms.assetid: 3714155e-b42f-8396-af66-6a1635f8631a
ms.date: 06/08/2017
---


# PublishObject.SourceType Property (PowerPoint)

Returns or sets the source type of the presentation to be published to HTML. Read/write.


## Syntax

 _expression_. **SourceType**

 _expression_ A variable that represents a **PublishObject** object.


### Return Value

PpPublishSourceType


## Remarks

The value of the  **SourceType** property can be one of these **PpPublishSourceType** constants.


||
|:-----|
|**ppPublishAll**|
|**ppPublishNamedSlideShow**|
|**ppPublishSlideRange**|
 Use the **ppPublishNamedSlideShow** value to publish a custom slide show, specifying the name of the custom slide show by using the **[SlideShowName](publishobject-slideshowname-property-powerpoint.md)** property.


## Example

This example publishes the specified slide range (slides three through five) of the active presentation to HTML. It names the published presentation Mallard.htm.


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

