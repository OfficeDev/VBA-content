---
title: ThreeDFormat.PresetMaterial Property (PowerPoint)
keywords: vbapp10.chm557014
f1_keywords:
- vbapp10.chm557014
ms.prod: powerpoint
api_name:
- PowerPoint.ThreeDFormat.PresetMaterial
ms.assetid: 71f224d4-6c2c-b42b-9a1a-a2ace4bb279f
ms.date: 06/08/2017
---


# ThreeDFormat.PresetMaterial Property (PowerPoint)

Returns or sets the extrusion surface material. Read/write.


## Syntax

 _expression_. **PresetMaterial**

 _expression_ A variable that represents a **ThreeDFormat** object.


### Return Value

MsoPresetMaterial


## Remarks

The value of the  **PresetMaterial** property can be one of these **MsoPresetMaterial** constants.


||
|:-----|
|**msoMaterialMatte**|
|**msoMaterialMetal**|
|**msoMaterialPlastic**|
|**msoMaterialWireFrame**|
|**msoPresetMaterialMixed**|

## Example

This example specifies that the extrusion surface for shape one in  `myDocument` be wire frame.


```vb
Set myDocument = ActivePresentation.Slides(1)

With myDocument.Shapes(1).ThreeD

    .Visible = True

    .PresetMaterial = msoMaterialWireFrame

End With
```


## See also


#### Concepts


[ThreeDFormat Object](threedformat-object-powerpoint.md)

