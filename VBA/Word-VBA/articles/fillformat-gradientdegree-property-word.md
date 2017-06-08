---
title: FillFormat.GradientDegree Property (Word)
keywords: vbawd10.chm164102247
f1_keywords:
- vbawd10.chm164102247
ms.prod: word
api_name:
- Word.FillFormat.GradientDegree
ms.assetid: c9fba9b0-cfbb-4cf1-c416-5886c77098fb
ms.date: 06/08/2017
---


# FillFormat.GradientDegree Property (Word)

Returns a value that indicates how dark or light a one-color gradient fill is. Read-only  **Single** .


## Syntax

 _expression_ . **GradientDegree**

 _expression_ A variable that represents a **[FillFormat](fillformat-object-word.md)** object.


## Remarks

A value of 0 (zero) means that black is mixed in with the shape's foreground color to form the gradient; a value of 1 means that white is mixed in; and values between 0 and 1 mean that a darker or lighter shade of the foreground color is mixed in. 

Use the  **[OneColorGradient](fillformat-onecolorgradient-method-word.md)** method to set the gradient degree for the fill.


## Example

This example adds a rectangle to the active document and sets the degree of its fill gradient to match that of the shape named "Rectangle 2." If Rectangle 2 doesn't have a one-color gradient fill, this example fails.


```vb
Dim docActive As Document 
Dim sngGradient As Single 
 
Set docActive = ActiveDocument 
With docActive.Shapes 
 sngGradient = .Item("Rectangle 2").Fill.GradientDegree 
 
 With .AddShape(msoShapeRectangle, 0, 0, 40, 80).Fill 
 .ForeColor.RGB = RGB(0, 128, 128) 
 .OneColorGradient msoGradientHorizontal, 1, sngGradient 
 End With 
End With
```


## See also


#### Concepts


[FillFormat Object](fillformat-object-word.md)

