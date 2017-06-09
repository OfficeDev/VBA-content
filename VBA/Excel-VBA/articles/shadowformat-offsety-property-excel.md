---
title: ShadowFormat.OffsetY Property (Excel)
keywords: vbaxl10.chm114005
f1_keywords:
- vbaxl10.chm114005
ms.prod: excel
api_name:
- Excel.ShadowFormat.OffsetY
ms.assetid: 54783d52-c32e-14ef-2cae-25f3a7676d80
ms.date: 06/08/2017
---


# ShadowFormat.OffsetY Property (Excel)

Returns or sets the vertical offset of the shadow from the specified shape, in points. A positive value offsets the shadow to the right of the shape; a negative value offsets it to the left. Read/write  **Single** .


## Syntax

 _expression_ . **OffsetY**

 _expression_ A variable that represents a **ShadowFormat** object.


## Remarks

If you want to nudge a shadow horizontally or vertically from its current position without having to specify an absolute position, use the  **[IncrementOffsetX](shadowformat-incrementoffsetx-method-excel.md)** method or the **[IncrementOffsetY](shadowformat-incrementoffsety-method-excel.md)** method.


## Example

This example sets the horizontal and vertical offsets for the shadow of shape three on  `myDocument`. The shadow is offset 5 points to the right of the shape and 3 points above it. If the shape doesn't already have a shadow, this example adds one to it.


```vb
Set myDocument = Worksheets(1) 
With myDocument.Shapes(3).Shadow 
 .Visible = True 
 .OffsetX = 5 
 .OffsetY = -3 
End With
```


## See also


#### Concepts


[ShadowFormat Object](shadowformat-object-excel.md)

