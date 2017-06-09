---
title: FillFormat.PresetGradientType Property (Word)
keywords: vbawd10.chm164102251
f1_keywords:
- vbawd10.chm164102251
ms.prod: word
api_name:
- Word.FillFormat.PresetGradientType
ms.assetid: b53ed5f8-61be-1abd-d3c7-e47a4ffc44b9
ms.date: 06/08/2017
---


# FillFormat.PresetGradientType Property (Word)

Returns the preset gradient type for the specified fill. Read-only  **MsoPresetGradientType** .


## Syntax

 _expression_ . **PresetGradientType**

 _expression_ An expression that represents a **[FillFormat](fillformat-object-word.md)** object.


## Remarks

Use the  **[PresetGradient](fillformat-presetgradient-method-word.md)** method to set the preset gradient type for the fill.


## Example

This example changes the fill for all shapes in  `myDocument` with the Moss preset gradient fill to the Fog preset gradient fill.


```vb
Set myDocument = ActiveDocument 
For Each s In myDocument.Shapes 
 With s.Fill 
 If .PresetGradientType = msoGradientMoss Then 
 .PresetGradient msoGradientHorizontal, 1, _ 
 msoGradientFog 
 End If 
 End With 
Next
```


## See also


#### Concepts


[FillFormat Object](fillformat-object-word.md)

