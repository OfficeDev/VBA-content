---
title: SlideRange.Duplicate Method (PowerPoint)
keywords: vbapp10.chm532015
f1_keywords:
- vbapp10.chm532015
ms.prod: powerpoint
api_name:
- PowerPoint.SlideRange.Duplicate
ms.assetid: 054b5be1-adbb-be83-1c25-e8585dbbdfe8
ms.date: 06/08/2017
---


# SlideRange.Duplicate Method (PowerPoint)

Creates a duplicate of the specified  **SlideRange** object, adds the new range of slides to the **Slides** collection immediately after the slide range specified originally, and then returns a **SlideRange** object that represents the duplicate slides.


## Syntax

 _expression_. **Duplicate**

 _expression_ A variable that represents a **SlideRange** object.


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


[SlideRange Object](sliderange-object-powerpoint.md)

