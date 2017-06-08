---
title: ShadowFormat.Obscured Property (Excel)
keywords: vbaxl10.chm114003
f1_keywords:
- vbaxl10.chm114003
ms.prod: excel
api_name:
- Excel.ShadowFormat.Obscured
ms.assetid: a2cc3324-d394-5332-41d2-e3733d0eb2d7
ms.date: 06/08/2017
---


# ShadowFormat.Obscured Property (Excel)

 **True** if the shadow of the specified shape appears filled in and is obscured by the shape, even if the shape has no fill. **False** if the shadow has no fill and the outline of the shadow is visible through the shape if the shape has no fill. Read/write **MsoTriState** .


## Syntax

 _expression_ . **Obscured**

 _expression_ A variable that represents a **ShadowFormat** object.


## Remarks





| **MsoTriState** can be one of these **MsoTriState** constants.|
| **msoCTrue**|
| **msoFalse** The shadow has no fill and the outline of the shadow is visible through the shape if the shape has no fill.|
| **msoTriStateMixed**|
| **msoTriStateToggle**|
| **msoTrue** The shadow of the specified shape appears filled in and is obscured by the shape, even if the shape has no fill.|

## Example

This example sets the horizontal and vertical offsets for the shadow of shape three on  `myDocument`. The shadow is offset 5 points to the right of the shape and 3 points above it. If the shape doesn't already have a shadow, this example adds one to it. The shadow will be filled in and obscured by the shape, even if the shape has no fill.


```vb
Set myDocument = Worksheets(1) 
With myDocument.Shapes(3).Shadow 
 .Visible = True 
 .OffsetX = 5 
 .OffsetY = -3 
 .Obscured = msoTrue 
End With
```


## See also


#### Concepts


[ShadowFormat Object](shadowformat-object-excel.md)

