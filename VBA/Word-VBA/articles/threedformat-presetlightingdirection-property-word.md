---
title: ThreeDFormat.PresetLightingDirection Property (Word)
keywords: vbawd10.chm164626537
f1_keywords:
- vbawd10.chm164626537
ms.prod: word
api_name:
- Word.ThreeDFormat.PresetLightingDirection
ms.assetid: 595b1541-c203-e736-2326-f7123f296d46
ms.date: 06/08/2017
---


# ThreeDFormat.PresetLightingDirection Property (Word)

Returns or sets the position of the light source relative to the extrusion. Read/write  **MsoPresetLightingDirection** .


## Syntax

 _expression_ . **PresetLightingDirection**

 _expression_ Required. A variable that represents a **[ThreeDFormat](threedformat-object-word.md)** object.


## Remarks

The lighting effects you set will not be apparent if the extrusion has a wireframe surface.


## Example

This example specifies that the extrusion for shape one on myDocument extend toward the top of the shape and that the lighting for the extrusion come from the left.


```vb
Set myDocument = ActiveDocument 
With myDocument.Shapes(1).ThreeD 
 .Visible = True 
 .SetExtrusionDirection msoExtrusionTop 
 .PresetLightingDirection = msoLightingLeft 
End With
```


## See also


#### Concepts


[ThreeDFormat Object](threedformat-object-word.md)

