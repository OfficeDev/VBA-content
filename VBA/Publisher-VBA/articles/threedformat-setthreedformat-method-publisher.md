---
title: ThreeDFormat.SetThreeDFormat Method (Publisher)
keywords: vbapb10.chm3801107
f1_keywords:
- vbapb10.chm3801107
ms.prod: publisher
api_name:
- Publisher.ThreeDFormat.SetThreeDFormat
ms.assetid: d73dbada-1a33-4b5b-9733-e228a0cc5f8c
ms.date: 06/08/2017
---


# ThreeDFormat.SetThreeDFormat Method (Publisher)

Sets the preset extrusion format. Each preset extrusion format contains a set of preset values for the 3-D properties of the extrusion.


## Syntax

 _expression_. **SetThreeDFormat**( **_PresetThreeDFormat_**)

 _expression_A variable that represents a  **ThreeDFormat** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
|PresetThreeDFormat|Required| **MsoPresetThreeDFormat**|Specifies a preset extrusion format that corresponds to one of the options (numbered from left to right, from top to bottom) displayed when you click the  **3-D** button on the **Drawing** toolbar.|

## Remarks

This method sets the  **[PresetThreeDFormat](threedformat-presetthreedformat-property-publisher.md)** property to the format specified by the PresetThreeDFormat argument.

The PresetThreeDFormat parameter can be one of the  **MsoPresetThreeDFormat** constants declared in the Microsoft Office type library and shown in the following table.



| **msoThreeD1**|
| **msoThreeD2**|
| **msoThreeD3**|
| **msoThreeD4**|
| **msoThreeD5**|
| **msoThreeD6**|
| **msoThreeD7**|
| **msoThreeD8**|
| **msoThreeD9**|
| **msoThreeD10**|
| **msoThreeD11**|
| **msoThreeD12**|
| **msoThreeD13**|
| **msoThreeD14**|
| **msoThreeD15**|
| **msoThreeD16**|
| **msoThreeD17**|
| **msoThreeD18**|
| **msoThreeD19**|
| **msoThreeD20**|

## Example

This example adds an oval to the active publication and sets its extrusion format to one of the preset 3-D formats.


```vb
With ActiveDocument.Pages(1).Shapes _ 
 .AddShape(Type:=msoShapeOval, _ 
 Left:=30, Top:=30, Width:=50, Height:=25).ThreeD 
 .Visible = True 
 .SetThreeDFormat PresetThreeDFormat:=msoThreeD12 
End With 

```


