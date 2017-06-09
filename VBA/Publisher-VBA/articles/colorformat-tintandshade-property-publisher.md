---
title: ColorFormat.TintAndShade Property (Publisher)
keywords: vbapb10.chm2555912
f1_keywords:
- vbapb10.chm2555912
ms.prod: publisher
api_name:
- Publisher.ColorFormat.TintAndShade
ms.assetid: 1c4897e0-ac55-08a8-8c43-dbd25d097ecc
ms.date: 06/08/2017
---


# ColorFormat.TintAndShade Property (Publisher)

Returns or sets a  **Single** that represents the lightening or darkening of a specified shape's color. Read/write.


## Syntax

 _expression_. **TintAndShade**

 _expression_A variable that represents a  **ColorFormat** object.


### Return Value

Single


## Remarks

You can enter a number from -1 (darkest) to 1 (lightest) for the  **TintAndShade** property, 0 (zero) being neutral.


## Example

This example creates a new shape in the active document, sets the fill color, and lightens the color shade.


```vb
Sub NewTintedShape() 
 Dim shpHeart As Shape 
 Set shpHeart = ActiveDocument.Pages(1).Shapes _ 
 .AddShape(Type:=msoShapeHeart, Left:=150, _ 
 Top:=150, Width:=250, Height:=250) 
 With shpHeart.Fill.ForeColor 
 .CMYK.SetCMYK Cyan:=255, Magenta:=28, Yellow:=0, Black:=0 
 .TintAndShade = 0.3 
 End With 
End Sub
```


