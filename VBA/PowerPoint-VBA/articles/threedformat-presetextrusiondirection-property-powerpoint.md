---
title: ThreeDFormat.PresetExtrusionDirection Property (PowerPoint)
keywords: vbapp10.chm557011
f1_keywords:
- vbapp10.chm557011
ms.prod: powerpoint
api_name:
- PowerPoint.ThreeDFormat.PresetExtrusionDirection
ms.assetid: 9bc0ba5b-c091-c385-3ef2-46994ed81347
ms.date: 06/08/2017
---


# ThreeDFormat.PresetExtrusionDirection Property (PowerPoint)

Returns the direction that the extrusion's sweep path takes away from the extruded shape (the front face of the extrusion). Read-only.


## Syntax

 _expression_. **PresetExtrusionDirection**

 _expression_ A variable that represents a **ThreeDFormat** object.


### Return Value

MsoPresetExtrusionDirection


## Remarks

This property is read-only. To set the value of this property, use the  **[SetExtrusionDirection](threedformat-setextrusiondirection-method-powerpoint.md)** method.

The value of the  **PresetExtrusionDirection** property can be one of these **MsoPresetExtrusionDirection** constants.


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

This example changes each extrusion on  `myDocument` that extends toward the upper-left corner of the extrusion's front face to an extrusion that extends toward the lower-right corner of the front face.


```vb
Set myDocument = ActivePresentation.Slides(1)

For Each s In myDocument.Shapes

    With s.ThreeD

        If .PresetExtrusionDirection = msoExtrusionTopLeft Then

            .SetExtrusionDirection msoExtrusionBottomRight

        End If

    End With

Next
```


## See also


#### Concepts


[ThreeDFormat Object](threedformat-object-powerpoint.md)

