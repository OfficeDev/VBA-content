---
title: PrintOptions.OutputType Property (PowerPoint)
keywords: vbapp10.chm517007
f1_keywords:
- vbapp10.chm517007
ms.prod: powerpoint
api_name:
- PowerPoint.PrintOptions.OutputType
ms.assetid: 673bcc73-bd60-13f9-f383-dd927401e0f6
ms.date: 06/08/2017
---


# PrintOptions.OutputType Property (PowerPoint)

Returns or sets a value that indicates which component (slides, handouts, notes pages, or an outline) of the presentation is to be printed. Read/write.


## Syntax

 _expression_. **OutputType**

 _expression_ A variable that represents an **PrintOptions** object.


### Return Value

PpPrintOutputType


## Remarks

The value of the  **OutputType** property can be one of these **PpPrintOutputType** constants.


||
|:-----|
|**ppPrintOutputBuildSlides**|
|**ppPrintOutputFourSlideHandouts**|
|**ppPrintOutputNineSlideHandouts**|
|**ppPrintOutputNotesPages**|
|**ppPrintOutputOneSlideHandouts**|
|**ppPrintOutputOutline**|
|**ppPrintOutputSixSlideHandouts**|
|**ppPrintOutputSlides**|
|**ppPrintOutputThreeSlideHandouts**|
|**ppPrintOutputTwoSlideHandouts**|

## Example

This example prints handouts of the active presentation with six slides to a page.


```vb
With ActivePresentation

    .PrintOptions.OutputType = ppPrintOutputSixSlideHandouts

    .PrintOut

End With
```


## See also


#### Concepts


[PrintOptions Object](printoptions-object-powerpoint.md)

