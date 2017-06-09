---
title: PlaySettings.StopAfterSlides Property (PowerPoint)
keywords: vbapp10.chm568009
f1_keywords:
- vbapp10.chm568009
ms.prod: powerpoint
api_name:
- PowerPoint.PlaySettings.StopAfterSlides
ms.assetid: 4c979acf-92b8-ebf6-16a3-ae9334dc4593
ms.date: 06/08/2017
---


# PlaySettings.StopAfterSlides Property (PowerPoint)

Returns or sets the number of slides to be displayed before the media clip stops playing. Read/write.


## Syntax

 _expression_. **StopAfterSlides**

 _expression_ A variable that represents a **PlaySettings** object.


### Return Value

Long


## Remarks

For the  **StopAfterSlides** property setting to take effect, the **[PauseAnimation](playsettings-pauseanimation-property-powerpoint.md)** property of the specified slide must be set to **False**, and the **[PlayOnEntry](playsettings-playonentry-property-powerpoint.md)** property must be set to **True**.

The media clip will stop playing when the specified number of slides have been displayed or when the clip comes to an end — whichever comes first. A value of 0 (zero) specifies that the clip will stop playing after the current slide.


## Example

This example specifies that the media clip represented by shape three on slide one in the active presentation will be played automatically when it is animated, that the slide show will continue while the media clip is playing in the background, and that the clip will stop playing after three slides are displayed or when the end of the clip is reached — whichever comes first. Shape three must be a sound or movie object.


```vb
Set OLEobj = ActivePresentation.Slides(1).Shapes(3)

With OLEobj.AnimationSettings.PlaySettings

    .PlayOnEntry = True

    .PauseAnimation = False

    .StopAfterSlides = 3

End With
```


## See also


#### Concepts


[PlaySettings Object](playsettings-object-powerpoint.md)

