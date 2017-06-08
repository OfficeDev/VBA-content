---
title: SlideShowTransition.Hidden Property (PowerPoint)
keywords: vbapp10.chm539007
f1_keywords:
- vbapp10.chm539007
ms.prod: powerpoint
api_name:
- PowerPoint.SlideShowTransition.Hidden
ms.assetid: 38e9add2-d05a-f0c3-6d8e-58e548d9789d
ms.date: 06/08/2017
---


# SlideShowTransition.Hidden Property (PowerPoint)

Determines whether the specified slide is hidden during a slide show. Read/write.


## Syntax

 _expression_. **Hidden**

 _expression_ A variable that represents a **SlideShowTransition** object.


### Return Value

MsoTriState


## Remarks

The value of the  **Hidden** property can be one of these **MsoTriState** constants.



|**Constant**|**Description**|
|:-----|:-----|
|**msoFalse**| The specified slide is not hidden during a slide show.|
|**msoTrue**| The specified slide is hidden during a slide show.|

## Example

This example makes slide two in the active presentation a hidden slide.


```vb
ActivePresentation.Slides(2).SlideShowTransition.Hidden = msoTrue
```


## See also


#### Concepts


[SlideShowTransition Object](slideshowtransition-object-powerpoint.md)

