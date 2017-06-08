---
title: ShapeRange.ParentGroup Property (PowerPoint)
keywords: vbapp10.chm548076
f1_keywords:
- vbapp10.chm548076
ms.prod: powerpoint
api_name:
- PowerPoint.ShapeRange.ParentGroup
ms.assetid: 425aec51-78d8-8e44-7d33-a300af184676
ms.date: 06/08/2017
---


# ShapeRange.ParentGroup Property (PowerPoint)

Returns a  **Shape** object that represents the common parent shape of a child shape or a range of child shapes.


## Syntax

 _expression_. **ParentGroup**

 _expression_ A variable that represents a **ShapeRange** object.


### Return Value

Shape


## Example

This example creates two shapes on the first slide in the active presentation and groups those shapes; then using one shape in the group, accesses the parent group and fills all shapes in the parent group with the same fill color. This example assumes that the first slide of the active presentation does not currently contain any shapes. If it does, you will receive an error.


```vb
Sub ParentGroup()

    Dim sldNewSlide As Slide
    Dim shpParentGroup As Shape

    'Add two shapes to active document and group

    Set sldNewSlide = ActivePresentation.Slides _
        .Add(Index:=1, Layout:=ppLayoutBlank)

    With sldNewSlide.Shapes
        .AddShape Type:=msoShapeBalloon, Left:=72, _
            Top:=72, Width:=100, Height:=100

        .AddShape Type:=msoShapeOval, Left:=110, _
            Top:=120, Width:=100, Height:=100

        .Range(Array(1, 2)).Group
    End With

    Set shpParentGroup = ActivePresentation.Slides(1).Shapes(1) _
        .GroupItems(1).ParentGroup

    shpParentGroup.Fill.ForeColor.RGB = RGB _
        (Red:=151, Green:=51, Blue:=250)

End Sub
```


## See also


#### Concepts


[ShapeRange Object](shaperange-object-powerpoint.md)

