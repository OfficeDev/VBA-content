---
title: TextEffectFormat Object (PowerPoint)
keywords: vbapp10.chm556000
f1_keywords:
- vbapp10.chm556000
ms.prod: powerpoint
api_name:
- PowerPoint.TextEffectFormat
ms.assetid: 62434479-237f-01c4-712c-08e48b391d48
ms.date: 06/08/2017
---


# TextEffectFormat Object (PowerPoint)

Contains properties and methods that apply to WordArt objects.


## Example

Use the  **TextEffect** property to return a **TextEffectFormat** object. The following example sets the font name and formatting for shape one on `myDocument`. For this example to work, shape one must be a WordArt object.


```vb
Set myDocument = ActivePresentation.Slides(1)

With myDocument.Shapes(1).TextEffect

    .FontName = "Courier New"

    .FontBold = True

    .FontItalic = True

End With
```


## See also


#### Concepts


[PowerPoint Object Model Reference](object-model-powerpoint-vba-reference.md)

