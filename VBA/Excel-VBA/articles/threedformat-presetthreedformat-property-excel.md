---
title: ThreeDFormat.PresetThreeDFormat Property (Excel)
keywords: vbaxl10.chm119013
f1_keywords:
- vbaxl10.chm119013
ms.prod: excel
api_name:
- Excel.ThreeDFormat.PresetThreeDFormat
ms.assetid: 678fa7f1-7cdc-ce05-98f7-bc6252eb3df1
ms.date: 06/08/2017
---


# ThreeDFormat.PresetThreeDFormat Property (Excel)

Returns the preset extrusion format. Each preset extrusion format contains a set of preset values for the various properties of the extrusion. Read-only  **MsoPresetThreeDFormat** .


## Syntax

 _expression_ . **PresetThreeDFormat**

 _expression_ A variable that represents a **ThreeDFormat** object.


## Remarks

If the extrusion has a custom format rather than a preset format, this property returns  **msoPresetThreeDFormatMixed** .



| **MsoPresetThreeDFormat** can be one of these **MsoPresetThreeDFormat** constants.|
| **msoThreeD1**|
| **msoThreeD11**|
| **msoThreeD13**|
| **msoThreeD15**|
| **msoThreeD17**|
| **msoThreeD19**|
| **msoThreeD20**|
| **msoThreeD4**|
| **msoThreeD6**|
| **msoThreeD8**|
| **msoPresetThreeDFormatMixed**|
| **msoThreeD10**|
| **msoThreeD12**|
| **msoThreeD14**|
| **msoThreeD16**|
| **msoThreeD18**|
| **msoThreeD2**|
| **msoThreeD3**|
| **msoThreeD5**|
| **msoThreeD7**|
| **msoThreeD9**|
This property is read-only. To set the preset extrusion format, use the  **[SetThreeDFormat](threedformat-setthreedformat-method-excel.md)** method.


## Example

This example sets the extrusion format for shape one on  `myDocument` to 3D Style 12 if the shape initially has a custom extrusion format.


```vb
Set myDocument = Worksheets(1) 
With myDocument.Shapes(1).ThreeD 
 If .PresetThreeDFormat = msoPresetThreeDFormatMixed Then 
 .SetThreeDFormat msoThreeD12 
 End If 
End With
```


## See also


#### Concepts


[ThreeDFormat Object](threedformat-object-excel.md)

