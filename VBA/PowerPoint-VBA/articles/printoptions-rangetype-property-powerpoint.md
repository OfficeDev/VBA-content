---
title: PrintOptions.RangeType Property (PowerPoint)
keywords: vbapp10.chm517011
f1_keywords:
- vbapp10.chm517011
ms.prod: powerpoint
api_name:
- PowerPoint.PrintOptions.RangeType
ms.assetid: 51d48974-16c9-0b96-9feb-651ca6347587
ms.date: 06/08/2017
---


# PrintOptions.RangeType Property (PowerPoint)

Returns or sets the type of print range for the presentation. Read/write.


## Syntax

 _expression_. **RangeType**

 _expression_ A variable that represents a **PrintOptions** object.


## Remarks

The value of the  **RangeType** property can be one of these **PpSlideShowRangeType** constants.


||
|:-----|
|**ppShowAll**|
|**ppShowNamedSlideShow**|
|**ppShowSlideRange**|
To print the slides ranges you've defined in the  **[PrintRanges](printranges-object-powerpoint.md)** collection, you must first set the **RangeType** property to **ppPrintSlideRange**. Setting **RangeType** to anything other than **ppPrintSlideRange** means that the ranges you've defined in the **PrintRanges** collection won't be applied. However, this doesn't affect the contents of the **PrintRanges** collection in any way. That is, if you define some print ranges, set the **RangeType** property to a value other than **ppPrintSlideRange**, and then later set **RangeType** back to **ppPrintSlideRange**, the print ranges you defined before will remain unchanged.

Specifying a value for the To and From arguments of the  **[PrintOut](presentation-printout-method-powerpoint.md)** method sets the value of this property.


## Example

This example prints the current slide the active presentation.


```vb
With ActivePresentation

    .PrintOptions.RangeType = ppPrintCurrent

    .PrintOut

End With
```


## See also


#### Concepts


[PrintOptions Object](printoptions-object-powerpoint.md)

