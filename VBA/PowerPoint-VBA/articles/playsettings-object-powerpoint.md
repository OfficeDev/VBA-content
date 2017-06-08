---
title: PlaySettings Object (PowerPoint)
keywords: vbapp10.chm568000
f1_keywords:
- vbapp10.chm568000
ms.prod: powerpoint
api_name:
- PowerPoint.PlaySettings
ms.assetid: 5a588b69-08ab-2422-12f9-a2666d3fc6ac
ms.date: 06/08/2017
---


# PlaySettings Object (PowerPoint)

Contains information about how the specified media clip will be played during a slide show.


## Example

Use the [PlaySettings](animationsettings-playsettings-property-powerpoint.md)property to return the  **PlaySettings** object. The following example inserts a movie named "Clock.avi" into slide one in the active presentation. It then sets it to be played automatically after the previous animation or slide transition, specifies that the slide show continue while the movie plays, and specifies that the movie object be hidden during a slide show except when it is playing.


```vb
Set clockMovie = ActivePresentation.Slides(1).Shapes _
    .AddMediaObject(FileName:="C:\WINNT\clock.avi", _
    Left:=20, Top:=20)
With clockMovie.AnimationSettings.PlaySettings
    .PlayOnEntry = True
    .PauseAnimation = False
    .HideWhileNotPlaying = True
End With
```


## See also


#### Concepts


[PowerPoint Object Model Reference](object-model-powerpoint-vba-reference.md)

