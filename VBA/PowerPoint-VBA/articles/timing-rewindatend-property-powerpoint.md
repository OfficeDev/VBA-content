---
title: Timing.RewindAtEnd Property (PowerPoint)
keywords: vbapp10.chm653015
f1_keywords:
- vbapp10.chm653015
ms.prod: powerpoint
api_name:
- PowerPoint.Timing.RewindAtEnd
ms.assetid: 2055f5aa-10d4-45a7-f25d-afaa924f0937
ms.date: 06/08/2017
---


# Timing.RewindAtEnd Property (PowerPoint)

Represents whether an object returns to its beginning position after an animation has ended. Read/write.


## Syntax

 _expression_. **RewindAtEnd**

 _expression_ A variable that represents a **Timing** object.


### Return Value

MsoTriState


## Remarks

The value of the  **RewindAtEnd** property can be one of these **MsoTriState** constants.



|**Constant**|**Description**|
|:-----|:-----|
|**msoFalse**| The object does not return to its beginning position after an animation has ended.|
|**msoTrue**| The object returns to its beginning position after an animation has ended.|

## Example

The following example adds a shape and an animation to the shape, then instructs the shape to return to its beginning position after the animation has ended.


```vb
Sub AddShapeSetTiming()

    Dim effDiamond As Effect
    Dim shpRectangle As Shape

    'Adds shape and sets animation effect

    Set shpRectangle = ActivePresentation.Slides(1).Shapes _
        .AddShape(Type:=msoShapeRectangle, Left:=100, _
        Top:=100, Width:=50, Height:=50)

    Set effDiamond = ActivePresentation.Slides(1).TimeLine.MainSequence _
        .AddEffect(Shape:=shpRectangle, effectId:=msoAnimEffectPathDiamond)

    'Sets duration of animation and returns shape to its original position

    With effDiamond.Timing
        .Duration = 3
        .RewindAtEnd = msoTrue
    End With

End Sub
```


## See also


#### Concepts


[Timing Object](timing-object-powerpoint.md)

