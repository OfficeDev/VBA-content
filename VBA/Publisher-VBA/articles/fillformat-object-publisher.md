---
title: FillFormat Object (Publisher)
keywords: vbapb10.chm2424831
f1_keywords:
- vbapb10.chm2424831
ms.prod: publisher
api_name:
- Publisher.FillFormat
ms.assetid: 0a5d4f7a-c42a-28ad-c86d-ac9828a3b874
ms.date: 06/08/2017
---


# FillFormat Object (Publisher)

Represents fill formatting for a shape. A shape can have a solid, gradient, texture, pattern, picture, or semitransparent fill.
 


## Remarks

Many of the properties of the  **FillFormat** object are read-only. To set one of these properties, you have to apply the corresponding method.
 

 

## Example

Use the  **[Fill](shape-fill-property-publisher.md)** property to return a **FillFormat** object. The following example adds a shape to the active document and then sets the gradient and color for the shape's fill.
 

 

```
Sub AddShapeAndSetFill() 
 With ActiveDocument.Pages(1).Shapes.AddShape(Type:=msoShapeHeart, _ 
 Left:=90, Top:=90, Width:=90, Height:=80).Fill 
 .ForeColor.RGB = RGB(Red:=255, Green:=0, Blue:=0) 
 .OneColorGradient Style:=msoGradientHorizontal, _ 
 Variant:=1, Degree:=1 
 End With 
End Sub
```


## Methods



|**Name**|
|:-----|
|[OneColorGradient](fillformat-onecolorgradient-method-publisher.md)|
|[Patterned](fillformat-patterned-method-publisher.md)|
|[PresetGradient](fillformat-presetgradient-method-publisher.md)|
|[PresetTextured](fillformat-presettextured-method-publisher.md)|
|[Solid](fillformat-solid-method-publisher.md)|
|[TwoColorGradient](fillformat-twocolorgradient-method-publisher.md)|
|[UserPicture](fillformat-userpicture-method-publisher.md)|
|[UserTextured](fillformat-usertextured-method-publisher.md)|

## Properties



|**Name**|
|:-----|
|[Application](fillformat-application-property-publisher.md)|
|[BackColor](fillformat-backcolor-property-publisher.md)|
|[ForeColor](fillformat-forecolor-property-publisher.md)|
|[GradientAngle](fillformat-gradientangle-property-publisher.md)|
|[GradientColorType](fillformat-gradientcolortype-property-publisher.md)|
|[GradientDegree](fillformat-gradientdegree-property-publisher.md)|
|[GradientStyle](fillformat-gradientstyle-property-publisher.md)|
|[GradientVariant](fillformat-gradientvariant-property-publisher.md)|
|[Parent](fillformat-parent-property-publisher.md)|
|[Pattern](fillformat-pattern-property-publisher.md)|
|[PresetGradientType](fillformat-presetgradienttype-property-publisher.md)|
|[PresetTexture](fillformat-presettexture-property-publisher.md)|
|[RotateWithObject](fillformat-rotatewithobject-property-publisher.md)|
|[TextureAlignment](fillformat-texturealignment-property-publisher.md)|
|[TextureHorizontalScale](fillformat-texturehorizontalscale-property-publisher.md)|
|[TextureName](fillformat-texturename-property-publisher.md)|
|[TextureOffsetX](fillformat-textureoffsetx-property-publisher.md)|
|[TextureOffsetY](fillformat-textureoffsety-property-publisher.md)|
|[TextureType](fillformat-texturetype-property-publisher.md)|
|[TextureVerticalScale](fillformat-textureverticalscale-property-publisher.md)|
|[Transparency](fillformat-transparency-property-publisher.md)|
|[Type](fillformat-type-property-publisher.md)|
|[Visible](fillformat-visible-property-publisher.md)|

