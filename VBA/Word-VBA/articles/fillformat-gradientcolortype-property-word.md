---
title: FillFormat.GradientColorType Property (Word)
keywords: vbawd10.chm164102246
f1_keywords:
- vbawd10.chm164102246
ms.prod: word
api_name:
- Word.FillFormat.GradientColorType
ms.assetid: 3722c4df-8091-6c66-b379-af8385ed9fc5
ms.date: 06/08/2017
---


# FillFormat.GradientColorType Property (Word)

Returns the gradient color type for the specified fill. Read-only  **MsoGradientColorType** .


## Syntax

 _expression_ . **GradientColorType**

 _expression_ An expression that represents a **[FillFormat](fillformat-object-word.md)** object.


## Remarks

This property is read-only. Use the  **[OneColorGradient](fillformat-onecolorgradient-method-word.md)** , **[PresetGradient](fillformat-presetgradient-method-word.md)** , or **[TwoColorGradient](fillformat-twocolorgradient-method-word.md)** method to set the gradient type for the fill.


## Example

This example changes the fill for all shapes in the active document that have a two-color gradient fill to a preset gradient fill.


```vb
Dim docActive As Document 
Dim shapeLoop As Shape 
 
Set docActive = ActiveDocument 
For Each shapeLoop In docActive.Shapes 
 With shapeLoop 
 .Fill 
 If .GradientColorType = msoGradientTwoColors Then 
 .PresetGradient msoGradientHorizontal, 1, _ 
 msoGradientBrass 
 End If 
 End With 
Next
```


## See also


#### Concepts


[FillFormat Object](fillformat-object-word.md)

