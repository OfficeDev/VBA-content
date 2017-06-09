---
title: ThreeDFormat Object (Publisher)
keywords: vbapb10.chm3866623
f1_keywords:
- vbapb10.chm3866623
ms.prod: publisher
api_name:
- Publisher.ThreeDFormat
ms.assetid: 11d57330-c99e-5aa9-d47c-2c5d2846ed4d
ms.date: 06/08/2017
---


# ThreeDFormat Object (Publisher)

Represents a shape's three-dimensional formatting.
 


## Remarks

You cannot apply three-dimensional formatting to some kinds of shapes, such as beveled shapes. Most of the properties and methods of the  **ThreeDFormat** object for such a shape will fail.
 

 

## Example

Use the  **[ThreeD](shape-threed-property-publisher.md)** property to return a **ThreeDFormat** object. This example sets the depth, extrusion color, extrusion direction, and lighting direction for the 3-D effects applied to shape one in the active publication.
 

 

```
Sub SetThreeDSettings() 
 Dim tdfTemp As ThreeDFormat 
 
 Set tdfTemp = _ 
 ActiveDocument.Pages(1).Shapes(1).ThreeD 
 
 With tdfTemp 
 .Visible = True 
 .Depth = 50 
 .ExtrusionColor.RGB = RGB(255, 100, 255) 
 .SetExtrusionDirection _ 
 PresetExtrusionDirection:=msoExtrusionTop 
 .PresetLightingDirection = msoLightingLeft 
 End With 
End Sub
```


## Methods



|**Name**|
|:-----|
|[IncrementRotationX](threedformat-incrementrotationx-method-publisher.md)|
|[IncrementRotationY](threedformat-incrementrotationy-method-publisher.md)|
|[ResetRotation](threedformat-resetrotation-method-publisher.md)|
|[SetExtrusionDirection](threedformat-setextrusiondirection-method-publisher.md)|
|[SetThreeDFormat](threedformat-setthreedformat-method-publisher.md)|

## Properties



|**Name**|
|:-----|
|[Application](threedformat-application-property-publisher.md)|
|[BevelBottomDepth](threedformat-bevelbottomdepth-property-publisher.md)|
|[BevelBottomInset](threedformat-bevelbottominset-property-publisher.md)|
|[BevelBottomType](threedformat-bevelbottomtype-property-publisher.md)|
|[BevelTopDepth](threedformat-beveltopdepth-property-publisher.md)|
|[BevelTopInset](threedformat-beveltopinset-property-publisher.md)|
|[BevelTopType](threedformat-beveltoptype-property-publisher.md)|
|[ContourColor](threedformat-contourcolor-property-publisher.md)|
|[ContourWidth](threedformat-contourwidth-property-publisher.md)|
|[Depth](threedformat-depth-property-publisher.md)|
|[ExtrusionColor](threedformat-extrusioncolor-property-publisher.md)|
|[ExtrusionColorType](threedformat-extrusioncolortype-property-publisher.md)|
|[FieldOfView](threedformat-fieldofview-property-publisher.md)|
|[Parent](threedformat-parent-property-publisher.md)|
|[Perspective](threedformat-perspective-property-publisher.md)|
|[PresetExtrusionDirection](threedformat-presetextrusiondirection-property-publisher.md)|
|[PresetLightingDirection](threedformat-presetlightingdirection-property-publisher.md)|
|[PresetLightingSoftness](threedformat-presetlightingsoftness-property-publisher.md)|
|[PresetMaterial](threedformat-presetmaterial-property-publisher.md)|
|[PresetThreeDFormat](threedformat-presetthreedformat-property-publisher.md)|
|[RotationX](threedformat-rotationx-property-publisher.md)|
|[RotationY](threedformat-rotationy-property-publisher.md)|
|[Visible](threedformat-visible-property-publisher.md)|

