---
title: TextEffectFormat.FontName Property (PowerPoint)
keywords: vbapp10.chm556006
f1_keywords:
- vbapp10.chm556006
ms.prod: powerpoint
api_name:
- PowerPoint.TextEffectFormat.FontName
ms.assetid: 4fdfc7a2-4b2e-e90f-719d-75a3f73c34e6
ms.date: 06/08/2017
---


# TextEffectFormat.FontName Property (PowerPoint)

Returns or sets the name of the font in the specified WordArt. Read/write.


## Syntax

 _expression_. **FontName**

 _expression_ A variable that represents a **TextEffectFormat** object.


### Return Value

String


## Example

This example sets the font name to "Courier New" for shape three on  `myDocument` if the shape is WordArt.


```vb
Set myDocument = ActivePresentation.Slides(1)

With myDocument.Shapes(3)

    If .Type = msoTextEffect Then

        .TextEffect.FontName = "Courier New"

    End If

End With
```


## See also


#### Concepts


[TextEffectFormat Object](texteffectformat-object-powerpoint.md)

