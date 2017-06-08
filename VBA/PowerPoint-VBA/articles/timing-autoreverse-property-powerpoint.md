---
title: Timing.AutoReverse Property (PowerPoint)
keywords: vbapp10.chm653012
f1_keywords:
- vbapp10.chm653012
ms.prod: powerpoint
api_name:
- PowerPoint.Timing.AutoReverse
ms.assetid: 82137189-a7f0-bacc-0550-41c9b5ff9ded
ms.date: 06/08/2017
---


# Timing.AutoReverse Property (PowerPoint)

Determines whether an effect should play forward and then in reverse, thereby doubling its duration. Read/write.


## Syntax

 _expression_. **AutoReverse**

 _expression_ A variable that represents an **Timing** object.


### Return Value

MsoTriState


## Remarks

The value of the  **AutoReverse** property can be one of these **MsoTriState** constants.



|**Constant**|**Description**|
|:-----|:-----|
|**msoFalse**| The default. The effect does not play forward and then in reverse.|
|**msoTrue**| The effect plays forward and then in reverse.|

## Example

The following example adds a shape and an animation effect to it; then it sets the animation to reverse direction after it finishes its forward movement.


```vb
Sub SetEffectTiming()

    Dim effDiamond As Effect
    Dim shpRectangle As Shape

    'Adds rectangle and applies diamond effect
    Set shpRectangle = ActivePresentation.Slides(1).Shapes.AddShape _
        (Type:=msoShapeRectangle, Left:=100, _
        Top:=100, Width:=50, Height:=50)

    Set effDiamond = ActivePresentation.Slides(1).TimeLine _
        .MainSequence.AddEffect(Shape:=shpRectangle, _
         effectId:=msoAnimEffectPathDiamond)

    'Sets the duration of and reverses the effect
    With effDiamond.Timing
        .Duration = 5 ' Length of effect.
        .AutoReverse = msoTrue
    End With

End Sub
```


## See also


#### Concepts


[Timing Object](timing-object-powerpoint.md)

