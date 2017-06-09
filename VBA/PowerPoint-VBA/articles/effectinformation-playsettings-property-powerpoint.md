---
title: EffectInformation.PlaySettings Property (PowerPoint)
keywords: vbapp10.chm655008
f1_keywords:
- vbapp10.chm655008
ms.prod: powerpoint
api_name:
- PowerPoint.EffectInformation.PlaySettings
ms.assetid: 702cf5b9-8164-cd25-e441-566a9a94fc14
ms.date: 06/08/2017
---


# EffectInformation.PlaySettings Property (PowerPoint)

Returns a  **[PlaySettings](playsettings-object-powerpoint.md)** object that contains information about how the specified media clip plays during a slide show. Read-only.


## Syntax

 _expression_. **PlaySettings**

 _expression_ A variable that represents an **EffectInformation** object.


### Return Value

PlaySettings


## Example

This example inserts a movie named Clock.avi onto slide one in the active presentation, sets it to play automatically after the slide transition, and specifies that the movie object be hidden during a slide show except when it is playing.


```vb
With ActivePresentation.Slides(1).Shapes.AddOLEObject(Left:=10, _
        Top:=10, Width:=250, Height:=250, _
    FileName:="c:\winnt\Clock.avi")
    With .AnimationSettings.PlaySettings
        .PlayOnEntry = True
        .HideWhileNotPlaying = True
    End With
End With
```


## See also


#### Concepts


[EffectInformation Object](effectinformation-object-powerpoint.md)


