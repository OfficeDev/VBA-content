---
title: ThreeDFormat.PresetThreeDFormat Property (PowerPoint)
keywords: vbapp10.chm557015
f1_keywords:
- vbapp10.chm557015
ms.prod: powerpoint
api_name:
- PowerPoint.ThreeDFormat.PresetThreeDFormat
ms.assetid: fcae7d2f-4d6d-6dfd-1693-fa46a85d1df2
ms.date: 06/08/2017
---


# ThreeDFormat.PresetThreeDFormat Property (PowerPoint)

Returns the preset extrusion format. Read-only.


## Syntax

 _expression_. **PresetThreeDFormat**

 _expression_ A variable that represents a **ThreeDFormat** object.


### Return Value

MsoPresetThreeDFormat


## Remarks

This property is read-only. To set the preset extrusion format, use the  **[SetThreeDFormat](threedformat-setthreedformat-method-powerpoint.md)** method.

Each preset extrusion format contains a set of preset values for the various properties of the extrusion. The values for this property correspond to the options (numbered from left to right, top to bottom) displayed when you click the  **3-D Rotation** submenu on the **Shape Effects** menu.

The value of the  **PresetThreeDFormat** property can be one of these **MsoPresetThreeDFormat** constants. If the value is **msoPresetThreeDFormatMixed**, the extrusion has a custom format rather than a preset format.


||
|:-----|
|**msoPresetThreeDFormatMixed**|
|**msoThreeD1**|
|**msoThreeD2**|
|**msoThreeD3**|
|**msoThreeD4**|
|**msoThreeD5**|
|**msoThreeD6**|
|**msoThreeD7**|
|**msoThreeD8**|
|**msoThreeD9**|
|**msoThreeD10**|
|**msoThreeD11**|
|**msoThreeD12**|
|**msoThreeD13**|
|**msoThreeD14**|
|**msoThreeD15**|
|**msoThreeD16**|
|**msoThreeD17**|
|**msoThreeD18**|
|**msoThreeD19**|
|**msoThreeD20**|

## Example

This example sets the extrusion format for shape one on  `myDocument` to 3D Style 12 if the shape initially has a custom extrusion format.


```vb
Set myDocument = ActivePresentation.Slides(1)

With myDocument.Shapes(1).ThreeD

    If .PresetThreeDFormat = msoPresetThreeDFormatMixed Then

        .SetThreeDFormat msoThreeD12

    End If

End With
```


## See also


#### Concepts


[ThreeDFormat Object](threedformat-object-powerpoint.md)

