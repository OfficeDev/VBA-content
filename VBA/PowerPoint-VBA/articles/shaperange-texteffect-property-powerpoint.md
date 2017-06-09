---
title: ShapeRange.TextEffect Property (PowerPoint)
keywords: vbapp10.chm548034
f1_keywords:
- vbapp10.chm548034
ms.prod: powerpoint
api_name:
- PowerPoint.ShapeRange.TextEffect
ms.assetid: 8cf70ead-8534-ef82-5064-21b9929e6f08
ms.date: 06/08/2017
---


# ShapeRange.TextEffect Property (PowerPoint)

Returns a  **[TextEffectFormat](texteffectformat-object-powerpoint.md)** object that contains text-effect formatting properties for the specified shape. Read-only.


## Syntax

 _expression_. **TextEffect**

 _expression_ A variable that represents a **ShapeRange** object.


### Return Value

TextEffectFormat


## Remarks

Applies to  **[ShapeRange](shaperange-object-powerpoint.md)** objects that represent WordArt.


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


[ShapeRange Object](shaperange-object-powerpoint.md)

