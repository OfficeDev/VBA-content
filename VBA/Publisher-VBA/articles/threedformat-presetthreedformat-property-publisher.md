---
title: ThreeDFormat.PresetThreeDFormat Property (Publisher)
keywords: vbapb10.chm3801352
f1_keywords:
- vbapb10.chm3801352
ms.prod: publisher
api_name:
- Publisher.ThreeDFormat.PresetThreeDFormat
ms.assetid: da0b2e3e-57e5-9c6f-6d08-3f60d38ba1f8
ms.date: 06/08/2017
---


# ThreeDFormat.PresetThreeDFormat Property (Publisher)

Returns an  **MsoPresetThreeDFormat** constant that represents the preset extrusion format. Read-only.


## Syntax

 _expression_. **PresetThreeDFormat**

 _expression_A variable that represents a  **ThreeDFormat** object.


### Return Value

MsoPresetThreeDFormat


## Remarks

The  **PresetThreeDFormat** property value can be one of the ** [MsoPresetThreeDFormat](http://msdn.microsoft.com/library/9d362115-1979-d079-d7e5-2e7788da614b%28Office.15%29.aspx)** constants declared in the Microsoft Office type library.

Each preset extrusion format contains a set of preset values for the various properties of the extrusion. If the extrusion has a custom format rather than a preset format, this property returns  **msoPresetThreeDFormatMixed**. 

The values for this property correspond to the options (numbered from left to right, top to bottom) displayed when you click the  **3-D Style** button on the **Formatting** toolbar.

Use the  **[SetThreeDFormat](threedformat-setthreedformat-method-publisher.md)** method to set the preset extrusion format.


## Example

This example sets the extrusion format for the first shape on the first page of the active publication to 3-D Style 12 if the shape initially has a custom extrusion format. For this example to work, the specified shape must be a 3-D shape.


```vb
Sub SetPreset3D() 
 With ActiveDocument.Pages(1).Shapes(1).ThreeD 
 If .PresetThreeDFormat = msoPresetThreeDFormatMixed Then 
 .SetThreeDFormat msoThreeD12 
 End If 
 End With 
End Sub
```


