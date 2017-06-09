---
title: Shape.ThreeD Property (Publisher)
keywords: vbapb10.chm2228305
f1_keywords:
- vbapb10.chm2228305
ms.prod: publisher
api_name:
- Publisher.Shape.ThreeD
ms.assetid: e3430bb2-2f2a-14a6-8eb4-98a29a96ad1c
ms.date: 06/08/2017
---


# Shape.ThreeD Property (Publisher)

Returns a  **[ThreeDFormat](threedformat-object-publisher.md)** object.


## Syntax

 _expression_. **ThreeD**

 _expression_A variable that represents a  **Shape** object.


## Remarks

Use the  **ThreeD** property to return a **ThreeDFormat** object whose properties are used to format the 3-D appearance of the specified shape.


## Example

This example sets the depth, extrusion color, extrusion direction, and lighting direction for the 3-D effects applied to shape one in the active publication.


```vb
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

```


