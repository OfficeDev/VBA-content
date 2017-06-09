---
title: TextEffectFormat.KernedPairs Property (PowerPoint)
keywords: vbapp10.chm556008
f1_keywords:
- vbapp10.chm556008
ms.prod: powerpoint
api_name:
- PowerPoint.TextEffectFormat.KernedPairs
ms.assetid: 03f0395e-ee31-80d2-7c0d-f404685a0e86
ms.date: 06/08/2017
---


# TextEffectFormat.KernedPairs Property (PowerPoint)

Determines whether the character pairs in the specified WordArt are kerned. Read/write.


## Syntax

 _expression_. **KernedPairs**

 _expression_ A variable that represents a **TextEffectFormat** object.


### Return Value

MsoTriState


## Remarks

The value returned by the  **KernedPairs** property can be one of these **MsoTriState** constants.



|**Constant**|**Description**|
|:-----|:-----|
|**msoFalse**|Character pairs in the specified WordArt are not kerned.|
|**msoTrue**| Character pairs in the specified WordArt are kerned.|

## Example

This example turns on character pair kerning for shape three on  `myDocument` if the shape is WordArt.


```vb
Set myDocument = ActivePresentation.Slides(1)

With myDocument.Shapes(3)

    If .Type = msoTextEffect Then

        .TextEffect.KernedPairs = msoTrue

    End If

End With
```


## See also


#### Concepts


[TextEffectFormat Object](texteffectformat-object-powerpoint.md)

