---
title: Shape.TextEffect Property (PowerPoint)
keywords: vbapp10.chm547034
f1_keywords:
- vbapp10.chm547034
ms.prod: powerpoint
api_name:
- PowerPoint.Shape.TextEffect
ms.assetid: b5d0a0a5-462d-1ede-3dac-7bedaaa1e318
ms.date: 06/08/2017
---


# Shape.TextEffect Property (PowerPoint)

Returns a  **[TextEffectFormat](texteffectformat-object-powerpoint.md)** object that contains text-effect formatting properties for the specified shape. Read-only.


## Syntax

 _expression_. **TextEffect**

 _expression_ A variable that represents a **Shape** object.


### Return Value

TextEffectFormat


## Remarks

Applies to  **[Shape](shape-object-powerpoint.md)** objects that represent WordArt.


## Example

This example sets the font style to bold for shape three on  `myDocument` if the shape is WordArt.


```vb
Set myDocument = ActivePresentation.Slides(1)

With myDocument.Shapes(3)

    If .Type = msoTextEffect Then

        .TextEffect.FontBold = True

    End If

End With
```


## See also


#### Concepts


[Shape Object](shape-object-powerpoint.md)

