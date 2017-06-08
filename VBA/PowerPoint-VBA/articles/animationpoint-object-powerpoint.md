---
title: AnimationPoint Object (PowerPoint)
keywords: vbapp10.chm664000
f1_keywords:
- vbapp10.chm664000
ms.prod: powerpoint
api_name:
- PowerPoint.AnimationPoint
ms.assetid: 79aa1a47-abab-f98f-955a-48be10a94c41
ms.date: 06/08/2017
---


# AnimationPoint Object (PowerPoint)

Represents an individual animation point for an animation behavior. The  **AnimationPoint** object is a member of the **[AnimationPoints](animationpoints-object-powerpoint.md)** collection. The **AnimationPoints** collection contains all the animation points for an animation behavior.


## Example

To add or reference an  **AnimationPoint** object, use the[Add](animationpoints-add-method-powerpoint.md)or [Item](animationpoints-item-method-powerpoint.md)method, respectively. Use the [Time](animationpoint-time-property-powerpoint.md)property of an  **AnimationPoint** object to set timing between animation points. Use the **[Value](animationpoint-value-property-powerpoint.md)** property to set other animation point properties, such as color. The following example adds three animation points to the first behavior in the active presentation's main animation sequence, and then it changes colors at each animation point.


```vb
Sub AniPoint()

    Dim sldNewSlide As Slide
    Dim shpHeart As Shape
    Dim effCustom As Effect
    Dim aniBehavior As AnimationBehavior
    Dim aptNewPoint As AnimationPoint

    Set sldNewSlide = ActivePresentation.Slides.Add _
        (Index:=1, Layout:=ppLayoutBlank)

    Set shpHeart = sldNewSlide.Shapes.AddShape _
        (Type:=msoShapeHeart, Left:=100, Top:=100, _
        Width:=200, Height:=200)

    Set effCustom = sldNewSlide.TimeLine.MainSequence _
        .AddEffect(shpHeart, msoAnimEffectCustom)

    Set aniBehavior = effCustom.Behaviors.Add(msoAnimTypeProperty)

    With aniBehavior.PropertyEffect
        .Property = msoAnimShapeFillColor
        Set aptNewPoint = .Points.Add
        aptNewPoint.Time = 0.2
        aptNewPoint.Value = RGB(0, 0, 0)
        Set aptNewPoint = .Points.Add
        aptNewPoint.Time = 0.5
        aptNewPoint.Value = RGB(0, 255, 0)
        Set aptNewPoint = .Points.Add
        aptNewPoint.Time = 1
        aptNewPoint.Value = RGB(0, 255, 255)
    End With

End Sub
```


## See also


#### Concepts


[PowerPoint Object Model Reference](object-model-powerpoint-vba-reference.md)

