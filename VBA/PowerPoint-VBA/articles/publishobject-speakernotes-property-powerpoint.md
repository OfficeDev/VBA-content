---
title: PublishObject.SpeakerNotes Property (PowerPoint)
keywords: vbapp10.chm635008
f1_keywords:
- vbapp10.chm635008
ms.prod: powerpoint
api_name:
- PowerPoint.PublishObject.SpeakerNotes
ms.assetid: 2dabb3db-4f94-c640-2c4d-d6c10551f903
ms.date: 06/08/2017
---


# PublishObject.SpeakerNotes Property (PowerPoint)

Determines whether speaker notes are to be published with the presentation. Read/write.


## Syntax

 _expression_. **SpeakerNotes**

 _expression_ A variable that represents a **PublishObject** object.


### Return Value

MsoTriState


## Remarks

The value of the  **SpeakerNotes** property can be one of these **MsoTriState** constants.



|**Constant**|**Description**|
|:-----|:-----|
|**msoFalse**|Speaker notes are not to be published with the presentation.|
|**msoTrue**| Speaker notes are to be published with the presentation.|

## Example

This example publishes slides three through five of the active presentation to HTML. It includes the associated speaker's notes with the published presentation and names it Mallard.htm.


```vb
With ActivePresentation.PublishObjects(1)

    .FileName = "C:\Test\Mallard.htm"

    .SourceType = ppPublishSlideRange

    .RangeStart = 3

    .RangeEnd = 5

    .SpeakerNotes = msoTrue

    .Publish

End With
```


## See also


#### Concepts


[PublishObject Object](publishobject-object-powerpoint.md)

