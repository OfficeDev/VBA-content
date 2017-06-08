---
title: Shape.TextEffect Property (Word)
keywords: vbawd10.chm161480824
f1_keywords:
- vbawd10.chm161480824
ms.prod: word
api_name:
- Word.Shape.TextEffect
ms.assetid: ce70ed2a-c448-cb12-db9f-a1f5d5d301e0
ms.date: 06/08/2017
---


# Shape.TextEffect Property (Word)

Returns a  **TextEffectFormat** object that contains text-effect formatting properties for the specified shape. Read-only.


## Syntax

 _expression_ . **TextEffect**

 _expression_ A variable that represents a **[Shape](shape-object-word.md)** object.


## Remarks

This property applies to  **Shape** objects that represent WordArt.


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


[Shape Object](shape-object-word.md)

