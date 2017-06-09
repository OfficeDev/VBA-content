---
title: ThreeDFormat.PresetLightingSoftness Property (Publisher)
keywords: vbapb10.chm3801350
f1_keywords:
- vbapb10.chm3801350
ms.prod: publisher
api_name:
- Publisher.ThreeDFormat.PresetLightingSoftness
ms.assetid: 8bad53c5-9d1c-367f-3f43-64691e193334
ms.date: 06/08/2017
---


# ThreeDFormat.PresetLightingSoftness Property (Publisher)

Returns or sets a  **MsoPresetLightingSoftness** constant that represents the intensity of the extrusion lighting. Read/write.


## Syntax

 _expression_. **PresetLightingSoftness**

 _expression_A variable that represents a  **ThreeDFormat** object.


### Return Value

MsoPresetLightingSoftness


## Remarks

The  **PresetLightingSoftness** property value can be one of the ** [MsoPresetLightingSoftness](http://msdn.microsoft.com/library/da5b4779-fca6-59cd-4cfe-599c3033c5d0%28Office.15%29.aspx)** constants declared in the Microsoft Office type library.


## Example

This example sets the extrusion for the first shape on the first page of the active publication to be lit brightly from the left. For this example to work, the specified shape must be a 3-D shape.


```vb
Sub SetExtrusionLightingBrighness() 
 With ActiveDocument.Pages(1).Shapes(1).ThreeD 
 .Visible = True 
 .PresetLightingSoftness = msoLightingBright 
 .PresetLightingDirection = msoLightingLeft 
 End With 
End Sub
```


