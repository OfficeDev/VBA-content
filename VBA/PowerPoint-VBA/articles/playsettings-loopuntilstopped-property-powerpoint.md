---
title: PlaySettings.LoopUntilStopped Property (PowerPoint)
keywords: vbapp10.chm568005
f1_keywords:
- vbapp10.chm568005
ms.prod: powerpoint
api_name:
- PowerPoint.PlaySettings.LoopUntilStopped
ms.assetid: b1c89b63-51cf-5ab3-4d98-2dd0a14f3d0e
ms.date: 06/08/2017
---


# PlaySettings.LoopUntilStopped Property (PowerPoint)

Determines whether the specified movie or sound loops continuously until either the next movie or sound starts, the user clicks the slide, or a slide transition occurs. Read/write.


## Syntax

 _expression_. **LoopUntilStopped**

 _expression_ A variable that represents a **PlaySettings** object.


## Remarks

The value of the  **LoopUntilStopped** property can be one of these **MsoTriState** constants.



|**Constant**|**Description**|
|:-----|:-----|
|**msoFalse**|The specified movie or sound does not loop continuously.|
|**msoTrue**| The specified movie or sound loops continuously until either the next movie or sound starts, the user clicks the slide, or a slide transition occurs.|

## Example

This example specifies that shape three on slide one in the active presentation will start to play in the animation order and will continue to play until the next media clip starts. Shape three must be a sound or movie object.


```vb
ActivePresentation.Slides(1).Shapes(3) _
    .AnimationSettings.PlaySettings.LoopUntilStopped = msoTrue
```


## See also


#### Concepts


[PlaySettings Object](playsettings-object-powerpoint.md)

