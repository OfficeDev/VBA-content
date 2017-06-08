---
title: FillFormat.GradientVariant Property (Publisher)
keywords: vbapb10.chm2359557
f1_keywords:
- vbapb10.chm2359557
ms.prod: publisher
api_name:
- Publisher.FillFormat.GradientVariant
ms.assetid: f57224dc-e9c6-e1aa-e950-97dfd5ed483f
ms.date: 06/08/2017
---


# FillFormat.GradientVariant Property (Publisher)

Returns a  **Long** indicating the gradient variant for the specified fill. Generally, values are integers from 1 to 4 for most gradient fills. If the gradient style is **msoGradientFromTitle** or **msoGradientFromCenter**, this property returns either 1 or 2. The values for this property correspond to the gradient variants (numbered from left to right and from top to bottom) on the  **Gradient** tab in the **Fill Effects** dialog box. Read-only.


## Syntax

 _expression_. **GradientVariant**

 _expression_A variable that represents a  **FillFormat** object.


### Return Value

Long


## Remarks

Use the  **[OneColorGradient](fillformat-onecolorgradient-method-publisher.md)**,  **[PresetGradient](fillformat-presetgradient-method-publisher.md)**, or  **[TwoColorGradient](fillformat-twocolorgradient-method-publisher.md)** method to set the gradient variant for the fill.


## Example

This example adds a rectangle to the active publication and sets its fill gradient variant to match that of the shape named rect1. For the example to work, rect1 must have a gradient fill.


```vb
Dim intVariant As Integer 
 
With ActiveDocument.Pages(1).Shapes 
 ' Store gradient variant of rect1. 
 intVariant = .Item("rect1").Fill.GradientVariant 
 ' Add new rectangle. 
 With .AddShape(Type:=msoShapeRectangle, _ 
 Left:=0, Top:=0, Width:=40, Height:=80).Fill 
 ' Set color and gradient of new rectangle. 
 .ForeColor.RGB = RGB(128, 0, 0) 
 .OneColorGradient Style:=msoGradientHorizontal, _ 
 Variant:=intVariant, Degree:=1 
 End With 
End With 

```


