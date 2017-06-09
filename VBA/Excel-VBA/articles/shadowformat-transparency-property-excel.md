---
title: ShadowFormat.Transparency Property (Excel)
keywords: vbaxl10.chm114006
f1_keywords:
- vbaxl10.chm114006
ms.prod: excel
api_name:
- Excel.ShadowFormat.Transparency
ms.assetid: c4a92bf1-3a29-16c0-0eb8-0b1faf75bd4a
ms.date: 06/08/2017
---


# ShadowFormat.Transparency Property (Excel)

Returns or sets the degree of transparency of the specified fill as a value from 0.0 (opaque) through 1.0 (clear). Read/write  **Double** .


## Syntax

 _expression_ . **Transparency**

 _expression_ A variable that represents a **ShadowFormat** object.


## Remarks

The value of this property affects the appearance of solid-colored fills and lines only; it has no effect on the appearance of patterned lines or patterned, gradient, picture, or textured fills.


## Example

This example sets the shadow of shape three on worksheet one to semitransparent red. If the shape doesn't already have a shadow, this example adds one to it.


```vb
With Worksheets(1).Shapes(3).Shadow 
 .Visible = True 
 .ForeColor.RGB = RGB(255, 0, 0) 
 .Transparency = 0.5 
End With
```


## See also


#### Concepts


[ShadowFormat Object](shadowformat-object-excel.md)

