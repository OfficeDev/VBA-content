---
title: AnimationPoints Object (PowerPoint)
keywords: vbapp10.chm663000
f1_keywords:
- vbapp10.chm663000
ms.prod: powerpoint
api_name:
- PowerPoint.AnimationPoints
ms.assetid: 6ea9ebc4-791c-9781-38c3-8b0973e0d152
ms.date: 06/08/2017
---


# AnimationPoints Object (PowerPoint)

Represents a collection of animation points for a  **[PropertyEffect](propertyeffect-object-powerpoint.md)** object.


## Example

Use the [Points](propertyeffect-points-property-powerpoint.md)property of the  **[PropertyEffect](propertyeffect-object-powerpoint.md)** object to return an **AnimationPoints** collection object. The following example adds an animation point to the first behavior in the active presentation's main animation sequence.


```vb
Sub AddPoint()
    ActivePresentation.Slides(1).TimeLine.MainSequence(1) _
        .Behaviors(1).PropertyEffect.Points.Add
End Sub
```

Transitions from one animation point to another can sometimes be abrupt or choppy. Use the [Smooth](animationpoints-smooth-property-powerpoint.md)property to make transitions smoother. This example smoothes the transitions between animation points.




```vb
Sub SmoothTransition()
    ActivePresentation.Slides(1).TimeLine.MainSequence(1) _
        .Behaviors(1).PropertyEffect.Points.Smooth = msoTrue
End Sub
```


## See also


#### Concepts


[PowerPoint Object Model Reference](object-model-powerpoint-vba-reference.md)

