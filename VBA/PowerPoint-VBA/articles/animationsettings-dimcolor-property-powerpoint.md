---
title: AnimationSettings.DimColor Property (PowerPoint)
keywords: vbapp10.chm565003
f1_keywords:
- vbapp10.chm565003
ms.prod: powerpoint
api_name:
- PowerPoint.AnimationSettings.DimColor
ms.assetid: 574c24b0-45af-2e7c-6fd5-bfc17f552c83
ms.date: 06/08/2017
---


# AnimationSettings.DimColor Property (PowerPoint)

Returns or sets a  **[ColorFormat](colorformat-object-powerpoint.md)** object that represents the color of the specified shape after it is been built. Read-only.


## Syntax

 _expression_. **DimColor**

 _expression_ A variable that represents a **AnimationSettings** object.


### Return Value

ColorFormat


## Remarks

If you don't get the effect you expect, check your other build settings. You won't see the effect of the  **DimColor** property unless the **[TextLevelEffect](animationsettings-textleveleffect-property-powerpoint.md)** property of the **AnimationSettings** object is set to something other than **ppAnimateLevelNone**, the **[AfterEffect](effectinformation-aftereffect-property-powerpoint.md)** property is set to **ppAfterEffectDim**, and the **[Animate](animationsettings-animate-property-powerpoint.md)** property is set to **True**. In addition, if the specified shape is the only item or the last item to be built on the slide, the shape won't be dimmed. To change the build order of the shapes on a slide, use the **[AnimationOrder](animationsettings-animationorder-property-powerpoint.md)** property.


## Example

This example adds a slide that contains both a title and a three-item list to the active presentation, sets the title and list to be dimmed after being built, and sets the color that each of them will be dimmed to.


```vb
With ActivePresentation.Slides.Add(2, ppLayoutText).Shapes
    With .Item(1)
        .TextFrame.TextRange.Text = "Sample title"
        With .AnimationSettings
            .TextLevelEffect = ppAnimateByAllLevels
            .AfterEffect = ppAfterEffectDim
            .DimColor.SchemeColor = ppShadow
            .Animate = True
        End With
    End With

    With .Item(2)
        .TextFrame.TextRange.Text = "Item one" _
            &; Chr(13) &; "Item two"
        With .AnimationSettings
            .TextLevelEffect = ppAnimateByFirstLevel
            .AfterEffect = ppAfterEffectDim
            .DimColor.RGB = RGB(100, 150, 130)
            .Animate = True
        End With
    End With
End With
```


## See also


#### Concepts


[AnimationSettings Object](animationsettings-object-powerpoint.md)

