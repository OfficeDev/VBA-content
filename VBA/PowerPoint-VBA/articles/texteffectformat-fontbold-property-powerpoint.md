---
title: TextEffectFormat.FontBold Property (PowerPoint)
keywords: vbapp10.chm556004
f1_keywords:
- vbapp10.chm556004
ms.prod: powerpoint
api_name:
- PowerPoint.TextEffectFormat.FontBold
ms.assetid: 3166f581-63f6-c2c1-1902-c6b3a511f244
ms.date: 06/08/2017
---


# TextEffectFormat.FontBold Property (PowerPoint)

Determines whether the font in the specified WordArt is bold. Read/write.


## Syntax

 _expression_. **FontBold**

 _expression_ A variable that represents a **TextEffectFormat** object.


### Return Value

MsoTriState


## Remarks

The value of the  **FontBold** property can be one of these **MsoTriState** constants.



|**Constant**|**Description**|
|:-----|:-----|
|**msoFalse**|The font in the specified WordArt is not bold.|
|**msoTrue**| The font in the specified WordArt is bold.|

## Example

This example sets the font to bold for shape three on  `myDocument` if the shape is WordArt.


```vb
Set myDocument = ActivePresentation.Slides(1)

With myDocument.Shapes(3)

    If .Type = msoTextEffect Then

        .TextEffect.FontBold = msoTrue

    End If

End With
```


## See also


#### Concepts


[TextEffectFormat Object](texteffectformat-object-powerpoint.md)

