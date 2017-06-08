---
title: FillFormat.GradientVariant Property (Word)
keywords: vbawd10.chm164102249
f1_keywords:
- vbawd10.chm164102249
ms.prod: word
api_name:
- Word.FillFormat.GradientVariant
ms.assetid: d92f56a2-fe56-4734-bddc-97517eea5def
ms.date: 06/08/2017
---


# FillFormat.GradientVariant Property (Word)

Returns the gradient variant for the specified fill as an integer value from 1 to 4 for most gradient fills. Read-only  **Long** .


## Syntax

 _expression_ . **GradientVariant**

 _expression_ A variable that represents a **[FillFormat](fillformat-object-word.md)** object.


## Remarks

If the gradient style is  **msoGradientFromCenter** , this property returns either 1 or 2. The values for this property correspond to the gradient variants (numbered from left to right and from top to bottom) on the **Gradient** tab in the **Fill Effects** dialog box.

Use the  **[OneColorGradient](fillformat-onecolorgradient-method-word.md)** or **[TwoColorGradient](fillformat-twocolorgradient-method-word.md)** method to set the gradient variant for the fill.


## Example

This example adds a rectangle to the active document and sets its fill gradient variant to match that of the shape named "rect1." For the example to work, rect1 must have a gradient fill.


```vb
Dim lngGradient As Long 
 
With ActiveDocument.Shapes 
 lngGradient = .Item("rect1").Fill.GradientVariant 
 With .AddShape(msoShapeRectangle, 0, 0, 40, 80).Fill 
 .ForeColor.RGB = RGB(128, 0, 0) 
 .OneColorGradient msoGradientHorizontal, _ 
 lngGradient, 1 
 End With 
End With
```


## See also


#### Concepts


[FillFormat Object](fillformat-object-word.md)

