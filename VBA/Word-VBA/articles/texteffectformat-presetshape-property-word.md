---
title: TextEffectFormat.PresetShape Property (Word)
keywords: vbawd10.chm164561003
f1_keywords:
- vbawd10.chm164561003
ms.prod: word
api_name:
- Word.TextEffectFormat.PresetShape
ms.assetid: 4d183208-7ea2-7179-4c6c-f710c16dd5fb
ms.date: 06/08/2017
---


# TextEffectFormat.PresetShape Property (Word)

Returns or sets the shape of the specified WordArt. Read/write  **MsoPresetTextEffectShape** .


## Syntax

 _expression_ . **PresetShape**

 _expression_ Required. A variable that represents a **[TextEffectFormat](texteffectformat-object-word.md)** object.


## Remarks

Setting the  **PresetTextEffect** property automatically sets the **PresetShape** property.


## Example

This example sets the shape of all WordArt on myDocument to a chevron whose center points down.


```vb
Set myDocument = ActiveDocument 
For Each s In myDocument.Shapes 
 If s.Type = msoTextEffect Then 
 s.TextEffect.PresetShape = msoTextEffectShapeChevronDown 
 End If 
Next
```


## See also


#### Concepts


[TextEffectFormat Object](texteffectformat-object-word.md)

