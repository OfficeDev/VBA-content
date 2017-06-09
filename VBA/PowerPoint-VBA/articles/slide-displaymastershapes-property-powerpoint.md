---
title: Slide.DisplayMasterShapes Property (PowerPoint)
keywords: vbapp10.chm531020
f1_keywords:
- vbapp10.chm531020
ms.prod: powerpoint
api_name:
- PowerPoint.Slide.DisplayMasterShapes
ms.assetid: 9a4a5146-e84d-b9fe-a837-0bcafa3fe61d
ms.date: 06/08/2017
---


# Slide.DisplayMasterShapes Property (PowerPoint)

Determines whether the specified slide displays the background objects on the slide master. Read/write.


## Syntax

 _expression_. **DisplayMasterShapes**

 _expression_ A variable that represents a **Slide** object.


### Return Value

MsoTriState


## Remarks

The value of the  **DisplayMasterShapes** property can be one of these **MsoTriState** constants.



|**Constant**|**Description**|
|:-----|:-----|
|**msoFalse**|The specified slide does not display the background objects on the slide master.|
|**msoTrue**| The specified slide displays the background objects on the slide master. These background objects can include text, drawings, OLE objects, and clip art you add to the slide master. Headers and footers aren't included.|
When you create a new slide, the default value for this property is  **msoTrue**. If you copy a slide from another presentation, it retains the setting it had in the original presentation. That is, if the slide omitted slide master background objects in the original presentation, it will omit them in the new presentation as well.

Note that the look of the slide's background is determined by the color scheme and background and by the background objects. If setting the  **DisplayMasterShapes** property alone doesn't give you the results you want, try setting the **FollowMasterBackground** and **ColorScheme** properties as well.


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


[Slide Object](slide-object-powerpoint.md)

