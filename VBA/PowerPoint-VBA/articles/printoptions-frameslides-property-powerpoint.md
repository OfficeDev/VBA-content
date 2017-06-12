---
title: PrintOptions.FrameSlides Property (PowerPoint)
keywords: vbapp10.chm517005
f1_keywords:
- vbapp10.chm517005
ms.prod: powerpoint
api_name:
- PowerPoint.PrintOptions.FrameSlides
ms.assetid: 4e866170-ab21-44e1-a497-7fc8e331fcec
ms.date: 06/08/2017
---


# PrintOptions.FrameSlides Property (PowerPoint)

Determines whether a thin frame is placed around the border of the printed slides. Read/write. 


## Syntax

 _expression_. **FrameSlides**

 _expression_ A variable that represents a **PrintOptions** object.


### Return Value

MsoTriState


## Remarks

The  **FrameSlides** property applies to printed slides, handouts, and notes pages.

The value of the  **FrameSlides** property can be one of these **MsoTriState** constants.



|**Constant**|**Description**|
|:-----|:-----|
|**msoFalse**|A thin frame is not placed around the border of the printed slides. |
|**msoTrue**| A thin frame is placed around the border of the printed slides.|

## Example

This example prints the active presentation with a frame around each slide.


```vb
With ActivePresentation

    .PrintOptions.FrameSlides = msoTrue

    .PrintOut

End With
```


## See also


#### Concepts


[PrintOptions Object](printoptions-object-powerpoint.md)

