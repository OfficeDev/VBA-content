---
title: SlideRange.Background Property (PowerPoint)
keywords: vbapp10.chm532007
f1_keywords:
- vbapp10.chm532007
ms.prod: powerpoint
api_name:
- PowerPoint.SlideRange.Background
ms.assetid: fdbda068-3038-b966-bf61-3527f0258ba4
ms.date: 06/08/2017
---


# SlideRange.Background Property (PowerPoint)

Returns a  **[ShapeRange](shaperange-object-powerpoint.md)** object that represents the slide background.


## Syntax

 _expression_. **Background**

 _expression_ A variable that represents a **SlideRange** object.


### Return Value

ShapeRange


## Remarks

If you use the  **Background** property to set the background for an individual slide without changing the slide master, the **FollowMasterBackground** property for that slide must be set to **False**.


## Example

This example sets the background of the slide master in the active presentation to a preset shade.


```vb
ActivePresentation.SlideMaster.Background.Fill.PresetGradient _
    Style:=msoGradientHorizontal, Variant:=1, _
    PresetGradientType:=msoGradientLateSunset
```

This example sets the background of slide one in the active presentation to a preset shade.




```vb
With ActivePresentation.Slides(1)
    .FollowMasterBackground = False
    .Background.Fill.PresetGradient Style:=msoGradientHorizontal, _
        Variant:=1, PresetGradientType:=msoGradientLateSunset
End With
```


## See also


#### Concepts


[SlideRange Object](sliderange-object-powerpoint.md)

