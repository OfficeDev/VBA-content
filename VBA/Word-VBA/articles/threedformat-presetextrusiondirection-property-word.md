---
title: ThreeDFormat.PresetExtrusionDirection Property (Word)
keywords: vbawd10.chm164626536
f1_keywords:
- vbawd10.chm164626536
ms.prod: word
api_name:
- Word.ThreeDFormat.PresetExtrusionDirection
ms.assetid: 8fc0cd0a-1d62-64ae-8757-851207aae56f
ms.date: 06/08/2017
---


# ThreeDFormat.PresetExtrusionDirection Property (Word)

Returns the direction taken by the extrusion's sweep path leading away from the extruded shape (the front face of the extrusion). Read/write  **MsoPresetExtrusionDirection** .


## Syntax

 _expression_ . **PresetExtrusionDirection**

 _expression_ Required. A variable that represents a **[ThreeDFormat](threedformat-object-word.md)** object.


## Remarks

This property is read-only. To set the value of this property, use the  **SetExtrusionDirection** method.


## Example

This example changes each extrusion on myDocument that extends toward the upper-left corner of the extrusion's front face to an extrusion that extends toward the lower-right corner of the front face.


```vb
Set myDocument = ActiveDocument 
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


[ThreeDFormat Object](threedformat-object-word.md)

