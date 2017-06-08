---
title: Slide.Master Property (PowerPoint)
keywords: vbapp10.chm531023
f1_keywords:
- vbapp10.chm531023
ms.prod: powerpoint
api_name:
- PowerPoint.Slide.Master
ms.assetid: cec5385d-f6af-dd8d-7989-251a70c4937e
ms.date: 06/08/2017
---


# Slide.Master Property (PowerPoint)

Returns a  **[Master](master-object-powerpoint.md)** object that represents the slide master. Read-only.


## Syntax

 _expression_. **Master**

 _expression_ A variable that represents a **Slide** object.


### Return Value

Master


## Example

This example sets the background fill for the slide master for slide one in the active presentation.


```vb
ActivePresentation.Slides(1).Master.Background.Fill _
    .PresetGradient msoGradientDiagonalUp, 1, msoGradientDaybreak
```


## See also


#### Concepts


[Slide Object](slide-object-powerpoint.md)

