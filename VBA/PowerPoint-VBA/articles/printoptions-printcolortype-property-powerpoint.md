---
title: PrintOptions.PrintColorType Property (PowerPoint)
keywords: vbapp10.chm517002
f1_keywords:
- vbapp10.chm517002
ms.prod: powerpoint
api_name:
- PowerPoint.PrintOptions.PrintColorType
ms.assetid: f552b2c6-fc25-4da9-c8e2-418c42e5df6c
ms.date: 06/08/2017
---


# PrintOptions.PrintColorType Property (PowerPoint)

Returns or sets the way the specified document will be printed: in black and white, in pure black and white (also referred to as high contrast), or in color. Read/write.


## Syntax

 _expression_. **PrintColorType**

 _expression_ A variable that represents a **PrintOptions** object.


### Return Value

PpPrintColorType


## Remarks

The value of the  **PrintColorType** property can be one of these **PpPrintColorType** constants. The default value is set by the printer.


||
|:-----|
|**ppPrintBlackAndWhite**|
|**ppPrintColor**|
|**ppPrintPureBlackAndWhite**|

## Example

This example prints the slides in the active presentation in color.


```vb
With Application.ActivePresentation

    .PrintOptions.PrintColorType = ppPrintColor

    .PrintOut

End With
```


## See also


#### Concepts


[PrintOptions Object](printoptions-object-powerpoint.md)

