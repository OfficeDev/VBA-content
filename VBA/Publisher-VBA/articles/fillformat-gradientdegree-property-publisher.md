---
title: FillFormat.GradientDegree Property (Publisher)
keywords: vbapb10.chm2359555
f1_keywords:
- vbapb10.chm2359555
ms.prod: publisher
api_name:
- Publisher.FillFormat.GradientDegree
ms.assetid: b2eba471-5f03-4904-f876-edea4d40a908
ms.date: 06/08/2017
---


# FillFormat.GradientDegree Property (Publisher)

Returns a  **Single** indicating how dark or light a one-color gradient fill is. A value of 0 (zero) means that black is mixed in with the shape's foreground color to form the gradient; a value of 1 means that white is mixed in; and values between 0 and 1 mean that a darker or lighter shade of the foreground color is mixed in. Read-only.


## Syntax

 _expression_. **GradientDegree**

 _expression_A variable that represents a  **FillFormat** object.


### Return Value

Single


## Remarks

Use the  **[OneColorGradient](fillformat-onecolorgradient-method-publisher.md)** method to set the gradient degree for the fill.


## Example

This example adds a rectangle to the active publication and sets the degree of its fill gradient to match that of the shape named Rectangle 2. If Rectangle 2 doesn't have a one-color gradient fill, this example generates an error.


```vb
Dim sngDegree As Single 
 
With ActiveDocument.Pages(1).Shapes 
 ' Store degree of one-color gradient. 
 sngDegree = .Item("Rectangle 2").Fill.GradientDegree 
 ' Add new rectangle. 
 With .AddShape(msoShapeRectangle, 0, 0, 40, 80).Fill 
 ' Set color and gradient for new rectangle. 
 .ForeColor.RGB = RGB(128, 0, 0) 
 .OneColorGradient Style:=msoGradientHorizontal, _ 
 Variant:=1, Degree:=sngDegree 
 End With 
End With 

```


