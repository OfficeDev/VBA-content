---
title: ThreeDFormat.SetThreeDFormat Method (Word)
keywords: vbawd10.chm164626445
f1_keywords:
- vbawd10.chm164626445
ms.prod: word
api_name:
- Word.ThreeDFormat.SetThreeDFormat
ms.assetid: 1fff9c23-0f40-ef9a-ca15-331caa61a107
ms.date: 06/08/2017
---


# ThreeDFormat.SetThreeDFormat Method (Word)

Sets the preset extrusion format.


## Syntax

 _expression_ . **SetThreeDFormat**( **_PresetThreeDFormat_** )

 _expression_ Required. A variable that represents a **[ThreeDFormat](threedformat-object-word.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _PresetThreeDFormat_|Required| **MsoPresetThreeDFormat**|Specifies a preset extrusion format that corresponds to one of the options (numbered from left to right, top to bottom) displayed when you click the  **3-D** button on the **Drawing** toolbar.|

## Remarks

Each preset extrusion format contains a set of preset values for the various properties of the extrusion. This method sets the  **PresetThreeDFormat** property to the format specified by the PresetThreeDFormat argument.


 **Note**  Specifying  **msoPresetThreeDFormatMixed** for the PresetThreeDFormat argument causes an error.


## Example

This example adds an oval to the active document and sets its extrusion format to 3D Style 12.


```vb
With ActiveDocument.Shapes.AddShape(msoShapeOval, _ 
 30, 30, 50, 25).ThreeD 
 .Visible = True 
 .SetThreeDFormat msoThreeD12 
End With
```


## See also


#### Concepts


[ThreeDFormat Object](threedformat-object-word.md)

