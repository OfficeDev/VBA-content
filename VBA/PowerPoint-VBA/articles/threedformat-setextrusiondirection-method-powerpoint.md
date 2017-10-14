---
title: ThreeDFormat.SetExtrusionDirection Method (PowerPoint)
keywords: vbapp10.chm557006
f1_keywords:
- vbapp10.chm557006
ms.prod: powerpoint
api_name:
- PowerPoint.ThreeDFormat.SetExtrusionDirection
ms.assetid: 3ce76681-1a37-258b-594c-11d1d4f161c6
ms.date: 06/08/2017
---


# ThreeDFormat.SetExtrusionDirection Method (PowerPoint)

Sets the direction that the extrusion's sweep path takes away from the extruded shape.


## Syntax

 _expression_. **SetExtrusionDirection**( **_PresetExtrusionDirection_** )

 _expression_ A variable that represents a **ThreeDFormat** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _PresetExtrusionDirection_|Required|**MsoPresetExtrusionDirection**|Specifies the extrusion direction.|

## Remarks

This method sets the  **[PresetExtrusionDirection](threedformat-presetextrusiondirection-property-powerpoint.md)** property to the direction specified by the PresetExtrusionDirection argument.

The PresetExtrusionDirection parameter value can be one of these  **MsoPresetExtrusionDirection** constants.


||
|:-----|
|**msoExtrusionBottom**|
|**msoExtrusionBottomLeft**|
|**msoExtrusionBottomRight**|
|**msoExtrusionLeft**|
|**msoExtrusionNone**|
|**msoExtrusionRight**|
|**msoExtrusionTop**|
|**msoExtrusionTopLeft**|
|**msoExtrusionTopRight**|
|**msoPresetExtrusionDirectionMixed**|

## Example

This example specifies that the extrusion for shape one on  `myDocument` extend toward the top of the shape and that the lighting for the extrusion come from the left.


```vb
Set myDocument = ActivePresentation.Slides(1)

With myDocument.Shapes(1).ThreeD

    .Visible = True

    .SetExtrusionDirection msoExtrusionTop

    .PresetLightingDirection = msoLightingLeft

End With
```


## See also


#### Concepts


[ThreeDFormat Object](threedformat-object-powerpoint.md)

