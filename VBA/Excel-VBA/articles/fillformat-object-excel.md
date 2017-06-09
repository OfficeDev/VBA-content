---
title: FillFormat Object (Excel)
keywords: vbaxl10.chm115000
f1_keywords:
- vbaxl10.chm115000
ms.prod: excel
api_name:
- Excel.FillFormat
ms.assetid: b602e09e-97ab-bfbe-1796-bc44ebb7dc28
ms.date: 06/08/2017
---


# FillFormat Object (Excel)

Represents fill formatting for a shape.


## Remarks

 A shape can have a solid, gradient, texture, pattern, picture, or semi-transparent fill.

Many of the properties of the  **FillFormat** object are read-only. To set one of these properties, you have to apply the corresponding method.


## Example

Use the  **[Fill](shape-fill-property-excel.md)** property to return a **FillFormat** object. The following example adds a rectangle to _myDocument_ and then sets the gradient and color for the rectangle's fill.


```
Set myDocument = Worksheets(1) 
With myDocument.Shapes.AddShape(msoShapeRectangle, _ 
 90, 90, 90, 80).Fill 
 .ForeColor.RGB = RGB(0, 128, 128) 
 .OneColorGradient msoGradientHorizontal, 1, 1 
End With
```


## Methods



|**Name**|
|:-----|
|[OneColorGradient](fillformat-onecolorgradient-method-excel.md)|
|[Patterned](fillformat-patterned-method-excel.md)|
|[PresetGradient](fillformat-presetgradient-method-excel.md)|
|[PresetTextured](fillformat-presettextured-method-excel.md)|
|[Solid](fillformat-solid-method-excel.md)|
|[TwoColorGradient](fillformat-twocolorgradient-method-excel.md)|
|[UserPicture](fillformat-userpicture-method-excel.md)|
|[UserTextured](fillformat-usertextured-method-excel.md)|

## Properties



|**Name**|
|:-----|
|[Application](fillformat-application-property-excel.md)|
|[BackColor](fillformat-backcolor-property-excel.md)|
|[Creator](fillformat-creator-property-excel.md)|
|[ForeColor](fillformat-forecolor-property-excel.md)|
|[GradientAngle](fillformat-gradientangle-property-excel.md)|
|[GradientColorType](fillformat-gradientcolortype-property-excel.md)|
|[GradientDegree](fillformat-gradientdegree-property-excel.md)|
|[GradientStops](fillformat-gradientstops-property-excel.md)|
|[GradientStyle](fillformat-gradientstyle-property-excel.md)|
|[GradientVariant](fillformat-gradientvariant-property-excel.md)|
|[Parent](fillformat-parent-property-excel.md)|
|[Pattern](fillformat-pattern-property-excel.md)|
|[PictureEffects](fillformat-pictureeffects-property-excel.md)|
|[PresetGradientType](fillformat-presetgradienttype-property-excel.md)|
|[PresetTexture](fillformat-presettexture-property-excel.md)|
|[RotateWithObject](fillformat-rotatewithobject-property-excel.md)|
|[TextureAlignment](fillformat-texturealignment-property-excel.md)|
|[TextureHorizontalScale](fillformat-texturehorizontalscale-property-excel.md)|
|[TextureName](fillformat-texturename-property-excel.md)|
|[TextureOffsetX](fillformat-textureoffsetx-property-excel.md)|
|[TextureOffsetY](fillformat-textureoffsety-property-excel.md)|
|[TextureTile](fillformat-texturetile-property-excel.md)|
|[TextureType](fillformat-texturetype-property-excel.md)|
|[TextureVerticalScale](fillformat-textureverticalscale-property-excel.md)|
|[Transparency](fillformat-transparency-property-excel.md)|
|[Type](fillformat-type-property-excel.md)|
|[Visible](fillformat-visible-property-excel.md)|

## See also


#### Other resources


[Excel Object Model Reference](http://msdn.microsoft.com/library/11ea8598-8a20-92d5-f98b-0da04263bf2c%28Office.15%29.aspx)
