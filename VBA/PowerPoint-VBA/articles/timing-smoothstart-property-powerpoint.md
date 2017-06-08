---
title: Timing.SmoothStart Property (PowerPoint)
keywords: vbapp10.chm653013
f1_keywords:
- vbapp10.chm653013
ms.prod: powerpoint
api_name:
- PowerPoint.Timing.SmoothStart
ms.assetid: 7e2f3578-7367-748d-7e3c-cd4643a71e9d
ms.date: 06/08/2017
---


# Timing.SmoothStart Property (PowerPoint)

Determines whether an animation should accelerate when it starts. Read/write.


## Syntax

 _expression_. **SmoothStart**

 _expression_ A variable that represents a **Timing** object.


### Return Value

MsoTriState


## Remarks

The value of the  **SmoothStart** property can be one of these **MsoTriState** constants.



|**Constant**|**Description**|
|:-----|:-----|
|**msoFalse**|The default. The animation does not accelerate when it starts. |
|**msoTrue**| The animation accelerates when it starts.|

## Example

The following example adds a shape to a slide, animates the shape, and instructs the shape to accelerate when it starts.


```vb
Sub AddShapeSetTiming()

    Dim effDiamond As Effect
    Dim shpRectangle As Shape

    Set shpRectangle = ActivePresentation.Slides(1).Shapes _
        .AddShape(Type:=msoShapeRectangle, Left:=100, _
        Top:=100, Width:=50, Height:=50)

    Set effDiamond = ActivePresentation.Slides(1).TimeLine.MainSequence _
        .AddEffect(Shape:=shpRectangle, effectId:=msoAnimEffectPathDiamond)

    With effDiamond.Timing
        .Duration = 5
        .SmoothStart = msoTrue
    End With

End Sub
```


## See also


#### Concepts


[Timing Object](timing-object-powerpoint.md)

