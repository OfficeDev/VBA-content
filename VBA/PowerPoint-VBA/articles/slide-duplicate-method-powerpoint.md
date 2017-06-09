---
title: Slide.Duplicate Method (PowerPoint)
keywords: vbapp10.chm531015
f1_keywords:
- vbapp10.chm531015
ms.prod: powerpoint
api_name:
- PowerPoint.Slide.Duplicate
ms.assetid: a098ddc4-9838-35f2-86c1-8d9e4ff40209
ms.date: 06/08/2017
---


# Slide.Duplicate Method (PowerPoint)

Creates a duplicate of the specified  **Slide** object, adds the new slide to the **Slides** collection immediately after the slide specified originally, and then returns a **Slide** object that represents the duplicate slide.


## Syntax

 _expression_. **Duplicate**

 _expression_ A variable that represents a **Slide** object.


### Return Value

SlideRange


## Example

This example creates a duplicate of slide one in the active presentation and then sets the background shading and the title text of the new slide. The new slide will be slide two in the presentation.


```vb
Set newSlide = ActivePresentation.Slides(1).Duplicate

With newSlide
    .Background.Fill.PresetGradient msoGradientVertical, _
        1, msoGradientGold
    .Shapes.Title.TextFrame.TextRange _
        .Text = "Second Quarter Earnings"
End With
```


## See also


#### Concepts


[Slide Object](slide-object-powerpoint.md)

