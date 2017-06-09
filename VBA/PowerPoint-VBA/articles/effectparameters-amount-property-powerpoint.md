---
title: EffectParameters.Amount Property (PowerPoint)
keywords: vbapp10.chm654004
f1_keywords:
- vbapp10.chm654004
ms.prod: powerpoint
api_name:
- PowerPoint.EffectParameters.Amount
ms.assetid: dcc17dbf-6064-64b1-5474-29918bc4e0c6
ms.date: 06/08/2017
---


# EffectParameters.Amount Property (PowerPoint)

Returns or sets a  **Single** that represents the number of degrees an animated shape is rotated around the z-axis. A positive value indicates clockwise rotation; a negative value indicates counterclockwise rotation. Read/write.


## Syntax

 _expression_. **Amount**

 _expression_ A variable that represents an **EffectParameters** object.


### Return Value

Single


## Example

The following example adds a shape, and a 90-degree spin animation to the shape.


```vb
Sub SetAnimEffect()

    Dim effSpin As Effect
    Dim shpCube As Shape

    Set shpCube = ActivePresentation.Slides(1).Shapes.AddShape _
        (Type:=msoShapeCube, Left:=100, Top:=100, _
        Width:=50, Height:=50)

    Set effSpin = ActivePresentation.Slides(1).TimeLine _
        .MainSequence.AddEffect(Shape:=shpCube, _
        effectId:=msoAnimEffectSpin)

    effSpin.Timing.Duration = 3
    effSpin.EffectParameters.Amount = -90

End Sub
```


## See also


#### Concepts


[EffectParameters Object](effectparameters-object-powerpoint.md)


