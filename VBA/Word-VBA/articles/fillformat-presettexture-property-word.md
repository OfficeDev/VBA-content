---
title: FillFormat.PresetTexture Property (Word)
keywords: vbawd10.chm164102252
f1_keywords:
- vbawd10.chm164102252
ms.prod: word
api_name:
- Word.FillFormat.PresetTexture
ms.assetid: 90503151-0351-26f3-de16-65cb21992f46
ms.date: 06/08/2017
---


# FillFormat.PresetTexture Property (Word)

Returns the preset texture for the specified fill. Read-only  **MsoPresetTexture** .


## Syntax

 _expression_ . **PresetTexture**

 _expression_ Required. An expression that returns a **[FillFormat](fillformat-object-word.md)** object.


## Remarks

Use the  **[PresetTextured](fillformat-presettextured-method-word.md)** method to specify the preset texture for the fill.


## Example

This example adds a rectangle to  `myDocument` and sets its preset texture to match that of shape two. For the example to work, shape two must have a preset textured fill.


```vb
Set myDocument = ActiveDocument 
With myDocument.Shapes 
 presetTexture2 = .Item(2).Fill.PresetTexture 
 .AddShape(msoShapeRectangle, 100, 0, 40, 80).Fill _ 
 .PresetTextured presetTexture2 
End With
```


## See also


#### Concepts


[FillFormat Object](fillformat-object-word.md)

