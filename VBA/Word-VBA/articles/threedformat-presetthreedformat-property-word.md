---
title: ThreeDFormat.PresetThreeDFormat Property (Word)
keywords: vbawd10.chm164626540
f1_keywords:
- vbawd10.chm164626540
ms.prod: word
api_name:
- Word.ThreeDFormat.PresetThreeDFormat
ms.assetid: 16a3b8d8-3fbf-670a-7d89-fac5f04a9512
ms.date: 06/08/2017
---


# ThreeDFormat.PresetThreeDFormat Property (Word)

Returns the preset extrusion format. Read-only  **MsoPresetThreeDFormat** .


## Syntax

 _expression_ . **PresetThreeDFormat**

 _expression_ Required. A variable that represents a **[ThreeDFormat](threedformat-object-word.md)** object.


## Remarks

Each preset extrusion format contains a set of preset values for the various properties of the extrusion. If the extrusion has a custom format rather than a preset format, this property returns  **msoPresetThreeDFormatMixed** .

The values for this property correspond to the options (numbered from left to right, top to bottom) displayed when you click the  **3-D** button on the **Drawing** toolbar.

Use the  **SetThreeDFormat** method to set the preset extrusion format.


## Example

This example sets the extrusion format for shape one on myDocument to 3-D Style 12 if the shape initially has a custom extrusion format.


```vb
Set myDocument = ActiveDocument 
With myDocument.Shapes(1).ThreeD 
 If .PresetThreeDFormat = msoPresetThreeDFormatMixed Then 
 .SetThreeDFormat msoThreeD12 
 End If 
End With
```


## See also


#### Concepts


[ThreeDFormat Object](threedformat-object-word.md)

