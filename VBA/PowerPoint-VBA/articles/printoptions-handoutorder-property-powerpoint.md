---
title: PrintOptions.HandoutOrder Property (PowerPoint)
keywords: vbapp10.chm517016
f1_keywords:
- vbapp10.chm517016
ms.prod: powerpoint
api_name:
- PowerPoint.PrintOptions.HandoutOrder
ms.assetid: d71782ef-42d6-6dd4-6812-3463d41e8173
ms.date: 06/08/2017
---


# PrintOptions.HandoutOrder Property (PowerPoint)

Returns or sets the page layout order in which slides appear on printed handouts that show multiple slides on one page. Read/write.


## Syntax

 _expression_. **HandoutOrder**

 _expression_ A variable that represents a **PrintOptions** object.


### Return Value

PpPrintHandoutOrder


## Remarks

The value of the  **HandoutOrder** property can be one of these **PpPrintHandoutOrder** constants.



|**Constant**|**Description**|
|:-----|:-----|
|**ppPrintHandoutHorizontalFirst**|Slides are ordered horizontally, with the first slide in the upper-left corner and the second slide to the right of it. If your language setting specifies a right-to-left language, the first slide is in the upper-right corner with the second slide to the left of it.|
|**ppPrintHandoutVerticalFirst**| Slides are ordered vertically, with the first slide in the upper-left corner and the second slide below it. If your language setting specifies a right-to-left language, the first slide is in the upper-right corner with the second slide below it.|

## Example

This example sets handouts of the active presentation to contain six slides per page, orders the slides horizontally on the handouts, and prints them.


```vb
With ActivePresentation

    .PrintOptions.OutputType = ppPrintOutputSixSlideHandouts

    .PrintOptions.HandoutOrder = ppPrintHandoutHorizontalFirst

    .PrintOut

End With
```


## See also


#### Concepts


[PrintOptions Object](printoptions-object-powerpoint.md)

