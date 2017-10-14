---
title: PrintOptions.SlideShowName Property (PowerPoint)
keywords: vbapp10.chm517014
f1_keywords:
- vbapp10.chm517014
ms.prod: powerpoint
api_name:
- PowerPoint.PrintOptions.SlideShowName
ms.assetid: 0eca6fce-23ba-0e13-a6a7-ee968f3aa973
ms.date: 06/08/2017
---


# PrintOptions.SlideShowName Property (PowerPoint)

Returns or sets the name of the custom slide show to print. Read/write .


## Syntax

 _expression_. **SlideShowName**

 _expression_ A variable that represents a **PrintOptions** object.


### Return Value

String


## Remarks

To print a custom slide show, you must first set the  **[RangeType](printoptions-rangetype-property-powerpoint.md)** property to **ppPrintNamedSlideShow**.


## Example

This example prints an existing custom slide show named "tech talk."


```vb
With ActivePresentation.PrintOptions

    .RangeType = ppPrintNamedSlideShow

    .SlideShowName = "tech talk"

End With

ActivePresentation.PrintOut
```


## See also


#### Concepts


[PrintOptions Object](printoptions-object-powerpoint.md)

