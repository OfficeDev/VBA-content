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



| <strong>msoThreeD1</strong>|
| 
<strong>msoThreeD2</strong>|
| 
<strong>msoThreeD3</strong>|
| 
<strong>msoThreeD4</strong>|
| 
<strong>msoThreeD5</strong>|
| 
<strong>msoThreeD6</strong>|
| 
<strong>msoThreeD7</strong>|
| 
<strong>msoThreeD8</strong>|
| 
<strong>msoThreeD9</strong>|
| 
<strong>msoThreeD10</strong>|
| 
<strong>msoThreeD11</strong>|
| 
<strong>msoThreeD12</strong>|
| 
<strong>msoThreeD13</strong>|
| 
<strong>msoThreeD14</strong>|
| 
<strong>msoThreeD15</strong>|
| 
<strong>msoThreeD16</strong>|
| 
<strong>msoThreeD17</strong>|
| 
<strong>msoThreeD18</strong>|
| 
<strong>msoThreeD19</strong>|
| 
<strong>msoThreeD20</strong>|

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


