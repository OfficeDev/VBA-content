---
title: Timing.Accelerate Property (PowerPoint)
keywords: vbapp10.chm653010
f1_keywords:
- vbapp10.chm653010
ms.prod: powerpoint
api_name:
- PowerPoint.Timing.Accelerate
ms.assetid: 3e1a7b53-e398-e814-56ed-9df19bb26a0d
ms.date: 06/08/2017
---


# Timing.Accelerate Property (PowerPoint)

Returns or sets the percentage of the duration over which a timing acceleration should take place. Read/write.


## Syntax

 _expression_. **Accelerate**

 _expression_ A variable that represents an **Timing** object.


### Return Value

Single


## Remarks

For example, a value of 0.9 means that an acceleration should start slower than the default speed for 90% of the total animation time, with the last 10% of the animation at the default speed. 

To slow down an animation at the end, use the  **[Decelerate](timing-decelerate-property-powerpoint.md)** property.


## Example

This example adds a shape and adds an animation, starting out slow and matching the default speed after 30% of the animation sequence.


```vb
Sub AddShapeSetTiming()

    Dim effDiamond As Effect
    Dim shpRectangle As Shape

    'Adds rectangle and specifies effect to use for rectangle
    Set shpRectangle = ActivePresentation.Slides(1) _
        .Shapes.AddShape(Type:=msoShapeRectangle, _
        Left:=100, Top:=100, Width:=50, Height:=50)

    Set effDiamond = ActivePresentation.Slides(1) _
        .TimeLine.MainSequence.AddEffect(Shape:=shpRectangle, _
        effectId:=msoAnimEffectPathDiamond)

    'Specifies the acceleration for the effect

    With effDiamond.Timing
        .Accelerate = 0.3
    End With

End Sub
```


## See also


#### Concepts


[Timing Object](timing-object-powerpoint.md)

