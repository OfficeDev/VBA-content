---
title: PropertyEffect Object (PowerPoint)
keywords: vbapp10.chm662000
f1_keywords:
- vbapp10.chm662000
ms.prod: powerpoint
api_name:
- PowerPoint.PropertyEffect
ms.assetid: 8fa129ac-f222-a01c-4545-0097b1164aee
ms.date: 06/08/2017
---


# PropertyEffect Object (PowerPoint)

Represents a property effect for an  **[AnimationBehavior](animationbehavior-object-powerpoint.md)** object.


## Example

Use the [PropertyEffect](animationbehavior-propertyeffect-property-powerpoint.md)property of the  **AnimationBehavior** object to return a **PropertyEffect** object. The following example refers to the property effect for a specified animation behavior.


```vb
ActivePresentation.Slides(1).TimeLine.MainSequence.Item(1) _
   .Behaviors(1).PropertyEffect
```

Use the  **[Points](propertyeffect-points-property-powerpoint.md)** property to access the animation points of a particular animation behavior. If you want to change only two states of an animation behavior, use the[From](propertyeffect-from-property-powerpoint.md)and [To](propertyeffect-to-property-powerpoint.md)properties. This example adds a new shape to the and sets the property effect to animate the fill color from blue to red.




```vb
Sub AddShapeSetAnimFill()

    Dim effBlinds As Effect
    Dim shpRectangle As Shape
    Dim animProperty As AnimationBehavior

    Set shpRectangle = ActivePresentation.Slides(1).Shapes _
        .AddShape(Type:=msoShapeRectangle, Left:=100, _
        Top:=100, Width:=50, Height:=50)

    Set effBlinds = ActivePresentation.Slides(1).TimeLine.MainSequence _
        .AddEffect(Shape:=shpRectangle, effectId:=msoAnimEffectBlinds)

    effBlinds.Timing.Duration = 3
    Set animProperty = effBlinds.Behaviors.Add(msoAnimTypeProperty)

    With animProperty.PropertyEffect
        .Property = msoAnimColor
        .From = RGB(Red:=0, Green:=0, Blue:=255)
        .To = RGB(Red:=255, Green:=0, Blue:=0)
    End With

End Sub
```


## See also


#### Concepts


[PowerPoint Object Model Reference](object-model-powerpoint-vba-reference.md)

