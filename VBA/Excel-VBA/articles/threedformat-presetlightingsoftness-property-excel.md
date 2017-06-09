---
title: ThreeDFormat.PresetLightingSoftness Property (Excel)
keywords: vbaxl10.chm119011
f1_keywords:
- vbaxl10.chm119011
ms.prod: excel
api_name:
- Excel.ThreeDFormat.PresetLightingSoftness
ms.assetid: e63a483b-16c6-edab-6a16-b539f0a424cb
ms.date: 06/08/2017
---


# ThreeDFormat.PresetLightingSoftness Property (Excel)

Returns or sets the intensity of the extrusion lighting. Read/write  **MsoPresetLightingSoftness** .


## Syntax

 _expression_ . **PresetLightingSoftness**

 _expression_ A variable that represents a **ThreeDFormat** object.


## Remarks





| **MsoPresetLightingSoftness** can be one of these **MsoPresetLightingSoftness** constants.|
| **msoLightingBright**|
| **msoLightingDim**|
| **msoLightingNormal**|
| **msoPresetLightingSoftnessMixed**|

## Example

This example specifies that the extrusion for shape one on  `myDocument` be lit brightly from the left.


```vb
Set myDocument = Worksheets(1) 
With myDocument.Shapes(1).ThreeD 
 .Visible = True 
 .PresetLightingSoftness = msoLightingBright 
 .PresetLightingDirection = msoLightingLeft 
End With
```


## See also


#### Concepts


[ThreeDFormat Object](threedformat-object-excel.md)

