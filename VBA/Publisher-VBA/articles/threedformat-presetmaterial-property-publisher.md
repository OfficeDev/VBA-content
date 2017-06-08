---
title: ThreeDFormat.PresetMaterial Property (Publisher)
keywords: vbapb10.chm3801351
f1_keywords:
- vbapb10.chm3801351
ms.prod: publisher
api_name:
- Publisher.ThreeDFormat.PresetMaterial
ms.assetid: 5f12fb22-f596-0d59-1f02-63ce8d4bd927
ms.date: 06/08/2017
---


# ThreeDFormat.PresetMaterial Property (Publisher)

Returns or sets an  **MsoPresetMaterial** constant that represents the extrusion surface material. Read/write.


## Syntax

 _expression_. **PresetMaterial**

 _expression_A variable that represents a  **ThreeDFormat** object.


### Return Value

MsoPresetMaterial


## Remarks

The  **PresetMaterial** property value can be one of the ** [MsoPresetMaterial](http://msdn.microsoft.com/library/4cf62ef4-f6c8-eb0c-1dfd-569aafca16c0%28Office.15%29.aspx)** constants declared in the Microsoft Office type library.


## Example

This example specifies that the extrusion surface for shape one in the active publication be a wireframe. For this example to work, the specified shape must be a 3-D shape.


```vb
Sub SetExtrusionMaterial() 
 With ActiveDocument.Pages(1).Shapes(1).ThreeD 
 .Visible = True 
 .PresetMaterial = msoMaterialWireFrame 
 End With 
End Sub
```


