---
title: PublishObject.Publish Method (PowerPoint)
keywords: vbapp10.chm635010
f1_keywords:
- vbapp10.chm635010
ms.prod: powerpoint
api_name:
- PowerPoint.PublishObject.Publish
ms.assetid: 890382ef-8aec-466d-40f9-e2bae6dc558b
ms.date: 06/08/2017
---


# PublishObject.Publish Method (PowerPoint)

Creates a Web presentation (HTML format) from any loaded presentation. You can view the published presentation in a Web browser.


## Syntax

 _expression_. **Publish**

 _expression_ A variable that represents a **PublishObject** object.


## Remarks

You can specify the content and attributes of the published presentation by setting various properties of the  **[PublishObject](publishobject-object-powerpoint.md)** object. For example, the **[SourceType](publishobject-sourcetype-property-powerpoint.md)** property defines the portion of a loaded presentation to be published. The **[RangeStart](publishobject-rangestart-property-powerpoint.md)** property and the **[RangeEnd](publishobject-rangeend-property-powerpoint.md)** property specify the range of slides to publish, and the **[SpeakerNotes](publishobject-speakernotes-property-powerpoint.md)** property designates whether or not to publish the speaker's notes.


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

