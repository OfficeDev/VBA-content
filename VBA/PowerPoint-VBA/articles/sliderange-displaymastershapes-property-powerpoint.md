---
title: SlideRange.DisplayMasterShapes Property (PowerPoint)
keywords: vbapp10.chm532020
f1_keywords:
- vbapp10.chm532020
ms.prod: powerpoint
api_name:
- PowerPoint.SlideRange.DisplayMasterShapes
ms.assetid: 1c30ec1d-4865-5fcd-12c5-70f3bfeffe7c
ms.date: 06/08/2017
---


# SlideRange.DisplayMasterShapes Property (PowerPoint)

Determines whether the specified range of slides displays the background objects on the slide master. Read/write.


## Syntax

 _expression_. **DisplayMasterShapes**

 _expression_ A variable that represents a **SlideRange** object.


### Return Value

MsoTriState


## Remarks

The value of the  **DisplayMasterShapes** property can be one of these **MsoTriState** constants.



|**Constant**|**Description**|
|:-----|:-----|
|**msoFalse**|The specified range of slides does not display the background objects on the slide master.|
|**msoTrue**| The specified range of slides displays the background objects on the slide master. These background objects can include text, drawings, OLE objects, and clip art you add to the slide master. Headers and footers aren't included.|

## Example

This example copies slide one from presentation two, pastes it at the end of presentation one, and matches the slide's background, color scheme, and background objects to the rest of presentation one.


```
Presentations(2).Slides(1).Copy

With Presentations(1).Slides.Paste

    .FollowMasterBackground = True

    .ColorScheme = Presentations(1).SlideMaster.ColorScheme

    .DisplayMasterShapes = msoTrue

End With
```


## See also


#### Concepts


[SlideRange Object](sliderange-object-powerpoint.md)

