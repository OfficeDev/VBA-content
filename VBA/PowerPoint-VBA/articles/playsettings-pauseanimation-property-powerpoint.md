---
title: PlaySettings.PauseAnimation Property (PowerPoint)
keywords: vbapp10.chm568008
f1_keywords:
- vbapp10.chm568008
ms.prod: powerpoint
api_name:
- PowerPoint.PlaySettings.PauseAnimation
ms.assetid: a27beaaa-9ed6-f7cf-8abe-9012d1337fa8
ms.date: 06/08/2017
---


# PlaySettings.PauseAnimation Property (PowerPoint)

Determines whether the slide show pauses until the specified media clip is finished playing. Read/write.


## Syntax

 _expression_. **PauseAnimation**

 _expression_ A variable that represents a **PlaySettings** object.


### Return Value

MsoTriState


## Remarks

For the  **PauseAnimation** property setting to take effect, the **[PlayOnEntry](playsettings-playonentry-property-powerpoint.md)** property of the specified shape must be set to **msoTrue**.

The value of the  **PauseAnimation** property can be one of these **MsoTriState** constants.



|**Constant**|**Description**|
|:-----|:-----|
|**msoFalse**|The slide show continues while the media clip plays in the background.|
|**msoTrue**| The slide show pauses until the specified media clip is finished playing.|

## Example

This example specifies that shape three on slide one in the active presentation will be played automatically when it is animated and that the slide show won't continue while the movie is playing in the background. Shape three must be a sound or movie object.


```vb
Set OLEobj = ActivePresentation.Slides(1).Shapes(3)

With OLEobj.AnimationSettings.PlaySettings

    .PlayOnEntry = msoTrue

    .PauseAnimation = msoTrue

End With
```


## See also


#### Concepts


[PlaySettings Object](playsettings-object-powerpoint.md)

