---
title: ThreeDFormat.PresetLightingDirection Property (Publisher)
keywords: vbapb10.chm3801349
f1_keywords:
- vbapb10.chm3801349
ms.prod: publisher
api_name:
- Publisher.ThreeDFormat.PresetLightingDirection
ms.assetid: 94957653-a4e1-bcb6-7697-ed10d1b54301
ms.date: 06/08/2017
---


# ThreeDFormat.PresetLightingDirection Property (Publisher)

Returns or sets an  **MsoPresetLightingDirection** constant that represents the position of the light source relative to the extrusion. Read/write.


## Syntax

 _expression_. **PresetLightingDirection**

 _expression_A variable that represents a  **ThreeDFormat** object.


### Return Value

MsoPresetLightingDirection


## Remarks

The  **PresetLightingDirection** property value can be one of the ** [MsoPresetLightingDirection](http://msdn.microsoft.com/library/d3de37f8-f4c8-d04f-12a9-5fb7340fb8b1%28Office.15%29.aspx)** constants declared in the Microsoft Office type library.

The lighting effects you set will not be apparent if the extrusion has a wireframe surface.


## Example

This example sets the extrusion for the first shape on the first page of the active publication to extend toward the top of the shape and that the lighting for the extrusion come from the left. For this example to work, the specified shape must be a 3-D shape.


```vb
Sub ExtrusionLighting() 
 With ActiveDocument.Pages(1).Shapes(1).ThreeD 
 .Visible = True 
 .SetExtrusionDirection msoExtrusionTop 
 .PresetLightingDirection = msoLightingLeft 
 End With 
End Sub
```


