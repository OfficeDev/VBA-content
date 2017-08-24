---
title: ThreeDFormat.SetExtrusionDirection Method (Publisher)
keywords: vbapb10.chm3801108
f1_keywords:
- vbapb10.chm3801108
ms.prod: publisher
api_name:
- Publisher.ThreeDFormat.SetExtrusionDirection
ms.assetid: ac01d31d-7775-8e33-3b68-6e53f952fdda
ms.date: 06/08/2017
---


# ThreeDFormat.SetExtrusionDirection Method (Publisher)

Sets the direction that the extrusion's sweep path takes away from the extruded shape.


## Syntax

 _expression_. **SetExtrusionDirection**( **_PresetExtrusionDirection_**)

 _expression_A variable that represents a  **ThreeDFormat** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
|PresetExtrusionDirection|Required| **MsoPresetExtrusionDirection**|Specifies the extrusion direction.|

## Remarks

The PresetExtrusionDirection parameter can be one of the  **MsoPresetExtrusionDirection** constants declared in the Microsoft Office type library and shown in the following table.



| **msoExtrusionBottom**|
| **msoExtrusionBottomLeft**|
| **msoExtrusionBottomRight**|
| **msoExtrusionLeft**|
| **msoExtrusionNone**|
| **msoExtrusionRight**|
| **msoExtrusionTop**|
| **msoExtrusionTopLeft**|
| **msoExtrusionTopRight**|
This method sets the  [PresetExtrusionDirection](threedformat-presetextrusiondirection-property-publisher.md)property to the direction specified by the PresetExtrusionDirection argument.


## Example

This example specifies that the extrusion for the first shape in the active publication extend toward the top of the shape and that the lighting for the extrusion come from the left.


```vb
With ActiveDocument.Pages(1).Shapes(1).ThreeD 
 .Visible = True 
 .SetExtrusionDirection _ 
 PresetExtrusionDirection:=msoExtrusionTop 
 .PresetLightingDirection = msoLightingLeft 
End With 

```


