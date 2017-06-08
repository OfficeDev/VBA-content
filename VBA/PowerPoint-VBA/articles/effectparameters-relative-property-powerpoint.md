---
title: EffectParameters.Relative Property (PowerPoint)
keywords: vbapp10.chm654007
f1_keywords:
- vbapp10.chm654007
ms.prod: powerpoint
api_name:
- PowerPoint.EffectParameters.Relative
ms.assetid: 2675d451-6123-d9df-8c83-a009037d5108
ms.date: 06/08/2017
---


# EffectParameters.Relative Property (PowerPoint)

Determines whether to set the motion position relative to the position of the shape. Read/write.


## Syntax

 _expression_. **Relative**

 _expression_ A variable that represents a **EffectParameters** object.


### Return Value

MsoTriState


## Remarks

This property is only used in conjunction with motion paths.

The value of the  **Relative** property can be one of these **MsoTriState** constants.



|**Constant**|**Description**|
|:-----|:-----|
|**msoFalse**|The default. The motion path is absolute.|
|**msoTrue**| The motion path is relative.|

## Example

The following example adds a shape, adds an animated motion path to the shape, and reports on its motion path relativity.


```vb
Sub AddShapeSetAnimPath()

    Dim effDiamond As Effect
    Dim shpCube As Shape

    Set shpCube = ActivePresentation.Slides(1).Shapes _
        .AddShape(Type:=msoShapeCube, Left:=100, _
        Top:=100, Width:=50, Height:=50)

    Set effDiamond = ActivePresentation.Slides(1).TimeLine.MainSequence _
        .AddEffect(Shape:=shpCube, effectId:=msoAnimEffectPathDiamond)

    effDiamond.Timing.Duration = 3

    MsgBox "Is motion path relative or absolute: " &; _
        effDiamond.EffectParameters.Relative &; vbCrLf &; _
        "0 = Relative, -1 = Absolute"
		
End Sub
```


## See also


#### Concepts



[EffectParameters Object](effectparameters-object-powerpoint.md)

