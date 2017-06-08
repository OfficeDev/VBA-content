---
title: PlaySettings.HideWhileNotPlaying Property (PowerPoint)
keywords: vbapp10.chm568004
f1_keywords:
- vbapp10.chm568004
ms.prod: powerpoint
api_name:
- PowerPoint.PlaySettings.HideWhileNotPlaying
ms.assetid: 04fb6933-b0ee-762a-f24b-662253647a16
ms.date: 06/08/2017
---


# PlaySettings.HideWhileNotPlaying Property (PowerPoint)

Determines whether the specified media clip is hidden during a slide show except when it is playing. Read/write.


## Syntax

 _expression_. **HideWhileNotPlaying**

 _expression_ A variable that represents a **PlaySettings** object.


### Return Value

MsoTriState


## Remarks

The value of the  **HideWhileNotPlaying** property can be one of these **MsoTriState** constants.



|**Constant**|**Description**|
|:-----|:-----|
|**msoFalse**|The specified media clip is not hidden during a slide show except when it is playing. |
|**msoTrue**| The specified media clip is hidden during a slide show except when it is playing.|

## Example

This example inserts a movie named "Clock.avi" onto slide one in the active presentation, sets it to play automatically after the slide transition, and specifies that the movie object be hidden during a slide show except when it is playing.


```vb
With ActivePresentation.Slides(1).Shapes _
        .AddOLEObject(Left:=10, Top:=10, _
        Width:=250, Height:=250, _
        FileName:="c:\winnt\clock.avi")

    With .AnimationSettings.PlaySettings
        .PlayOnEntry = True
        .HideWhileNotPlaying = msoTrue
    End With

End With
```


## See also


#### Concepts


[PlaySettings Object](playsettings-object-powerpoint.md)

