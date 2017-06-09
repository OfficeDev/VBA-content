---
title: ThreeDFormat.PresetExtrusionDirection Property (Publisher)
keywords: vbapb10.chm3801348
f1_keywords:
- vbapb10.chm3801348
ms.prod: publisher
api_name:
- Publisher.ThreeDFormat.PresetExtrusionDirection
ms.assetid: fdf3843e-12bc-4b3b-11cb-e512abd991af
ms.date: 06/08/2017
---


# ThreeDFormat.PresetExtrusionDirection Property (Publisher)

Returns an  **MsoPresetExtrusionDirection** constant that represents the direction taken by the extrusion's sweep path leading away from the extruded shape (the front face of the extrusion). Read-only.


## Syntax

 _expression_. **PresetExtrusionDirection**

 _expression_A variable that represents a  **ThreeDFormat** object.


### Return Value

MsoPresetExtrusionDirection


## Remarks

The  **PresetExtrusionDirection** property value can be one of the **MsoPresetExtrusionDirection** constants declared in the Microsoft Office type library and shown in the following table.



| **msoExtrusionBottom**|
| **msoExtrusionBottomLeft**|
| **msoExtrusionBottomRight**|
| **msoExtrusionLeft**|
| **msoExtrusionNone**|
| **msoExtrusionRight**|
| **msoExtrusionTop**|
| **msoExtrusionTopLeft**|
| **msoExtrusionTopRight**|
| **msoPresetExtrusionDirectionMixed**|
This property is read-only. To set the value of this property, use the  **[SetExtrusionDirection](threedformat-setextrusiondirection-method-publisher.md)** method.


## Example

This example changes the extrusion for the first shape on the first page of the active publication if the extrusion extends toward the upper-left corner of the extrusion's front face. For this example to work, the specified shape must be a 3-D shape.


```vb
Sub SetExtrusion() 
 With ActiveDocument.Pages(1).Shapes(1).ThreeD 
 If .PresetExtrusionDirection = msoExtrusionTopLeft Then 
 .SetExtrusionDirection msoExtrusionBottomRight 
 End If 
 End With 
End Sub
```


