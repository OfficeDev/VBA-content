---
title: FillFormat.Background Method (PowerPoint)
keywords: vbapp10.chm552002
f1_keywords:
- vbapp10.chm552002
ms.prod: powerpoint
api_name:
- PowerPoint.FillFormat.Background
ms.assetid: 4c82e3d3-86cd-d18f-ead1-9fc2dda5efd8
ms.date: 06/08/2017
---


# FillFormat.Background Method (PowerPoint)

Specifies that the shape's fill should match the slide background. If you change the slide background after applying this method to a fill, the fill will also change.


## Syntax

 _expression_. **Background**

 _expression_ A variable that represents a **FillFormat** object.


## Remarks

Note that applying the  **Background** method to a shape's fill isn't the same as setting a transparent fill for the shape, nor is it always the same as applying the same fill to the shape as you apply to the background. The second example demonstrates this.


## Example

This example sets the fill of shape one on slide one in the active presentation to match the slide background.


```vb
ActivePresentation.Slides(1).Shapes(1).Fill.Background
```

This example sets the background for slide one in the active presentation to a preset gradient, adds a rectangle to the slide, and then places three ovals in front of the rectangle. The first oval has a fill that matches the slide background, the second has a transparent fill, and the third has the same fill applied to it as was applied to the background. Notice the difference in the appearances of these three ovals.




```vb
With ActivePresentation.Slides(1)

    .FollowMasterBackground = False
    .Background.Fill.PresetGradient _
        msoGradientHorizontal, 1, msoGradientDaybreak

    With .Shapes
        .AddShape msoShapeRectangle, 50, 200, 600, 100
        .AddShape(msoShapeOval, 75, 150, 150, 100) _
            .Fill.Background
        .AddShape(msoShapeOval, 275, 150, 150, 100).Fill _
            .Transparency = 1
        .AddShape(msoShapeOval, 475, 150, 150, 100) _
            .Fill.PresetGradient _
            msoGradientHorizontal, 1, msoGradientDaybreak
    End With
	
End With
```


## See also


#### Concepts


[FillFormat Object](fillformat-object-powerpoint.md)

