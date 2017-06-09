---
title: InlineShape.TextEffect Property (Word)
keywords: vbawd10.chm162005112
f1_keywords:
- vbawd10.chm162005112
ms.prod: word
api_name:
- Word.InlineShape.TextEffect
ms.assetid: 349563af-6a14-a8d9-c0a4-829910d7dc2c
ms.date: 06/08/2017
---


# InlineShape.TextEffect Property (Word)

Returns a  **TextEffectFormat** object that contains text-effect formatting properties for the specified inline shape. Read-only.


## Syntax

 _expression_ . **TextEffect**

 _expression_ A variable that represents an **[InlineShape](inlineshape-object-word.md)** object.


## Example

This example sets the font style to bold for shape three on  _myDocument_ if the shape is WordArt.


```vb
Set myDocument = ActiveDocument 
With myDocument.Shapes(3) 
 If .Type = msoTextEffect Then 
 .TextEffect.FontBold = True 
 End If 
End With
```


## See also


#### Concepts


[InlineShape Object](inlineshape-object-word.md)

