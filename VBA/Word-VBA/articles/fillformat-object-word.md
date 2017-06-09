---
title: FillFormat Object (Word)
keywords: vbawd10.chm2504
f1_keywords:
- vbawd10.chm2504
ms.prod: word
api_name:
- Word.FillFormat
ms.assetid: 39205d07-9e37-1be1-ec4a-93ba8bac2f26
ms.date: 06/08/2017
---


# FillFormat Object (Word)

Represents fill formatting for a shape. A shape can have a solid, gradient, texture, pattern, picture, or semi-transparent fill.


## Remarks

Use the  **Fill** property to return a **FillFormat** object. The following example adds a rectangle to the active document and then sets the gradient and color for the rectangle's fill.


```
With ActiveDocument.Shapes _ 
 .AddShape(msoShapeRectangle, 90, 90, 90, 80).Fill 
 .ForeColor.RGB = RGB(0, 128, 128) 
 .OneColorGradient msoGradientHorizontal, 1, 1 
End With
```

Many of the properties of the  **FillFormat** object are read-only. To set one of these properties, you have to apply the corresponding method.


## Methods



|**Name**|
|:-----|
|[OneColorGradient](fillformat-onecolorgradient-method-word.md)|
|[Patterned](fillformat-patterned-method-word.md)|
|[PresetGradient](fillformat-presetgradient-method-word.md)|
|[PresetTextured](fillformat-presettextured-method-word.md)|
|[Solid](fillformat-solid-method-word.md)|
|[TwoColorGradient](fillformat-twocolorgradient-method-word.md)|
|[UserPicture](fillformat-userpicture-method-word.md)|
|[UserTextured](fillformat-usertextured-method-word.md)|

## Properties



|**Name**|
|:-----|
|[Application](fillformat-application-property-word.md)|
|[BackColor](fillformat-backcolor-property-word.md)|
|[Creator](fillformat-creator-property-word.md)|
|[ForeColor](fillformat-forecolor-property-word.md)|
|[GradientAngle](fillformat-gradientangle-property-word.md)|
|[GradientColorType](fillformat-gradientcolortype-property-word.md)|
|[GradientDegree](fillformat-gradientdegree-property-word.md)|
|[GradientStops](fillformat-gradientstops-property-word.md)|
|[GradientStyle](fillformat-gradientstyle-property-word.md)|
|[GradientVariant](fillformat-gradientvariant-property-word.md)|
|[Parent](fillformat-parent-property-word.md)|
|[Pattern](fillformat-pattern-property-word.md)|
|[PictureEffects](fillformat-pictureeffects-property-word.md)|
|[PresetGradientType](fillformat-presetgradienttype-property-word.md)|
|[PresetTexture](fillformat-presettexture-property-word.md)|
|[RotateWithObject](fillformat-rotatewithobject-property-word.md)|
|[TextureAlignment](fillformat-texturealignment-property-word.md)|
|[TextureHorizontalScale](fillformat-texturehorizontalscale-property-word.md)|
|[TextureName](fillformat-texturename-property-word.md)|
|[TextureOffsetX](fillformat-textureoffsetx-property-word.md)|
|[TextureOffsetY](fillformat-textureoffsety-property-word.md)|
|[TextureTile](fillformat-texturetile-property-word.md)|
|[TextureType](fillformat-texturetype-property-word.md)|
|[TextureVerticalScale](fillformat-textureverticalscale-property-word.md)|
|[Transparency](fillformat-transparency-property-word.md)|
|[Type](fillformat-type-property-word.md)|
|[Visible](fillformat-visible-property-word.md)|

## See also


#### Other resources


[Word Object Model Reference](http://msdn.microsoft.com/library/be452561-b436-bb9b-6f94-3faa9a74a6fd%28Office.15%29.aspx)
