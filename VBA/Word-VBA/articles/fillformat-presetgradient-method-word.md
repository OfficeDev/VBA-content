---
title: FillFormat.PresetGradient Method (Word)
keywords: vbawd10.chm164102157
f1_keywords:
- vbawd10.chm164102157
ms.prod: word
api_name:
- Word.FillFormat.PresetGradient
ms.assetid: bffe754d-6593-9684-abf4-b5d1e9df720e
ms.date: 06/08/2017
---


# FillFormat.PresetGradient Method (Word)

Sets the specified fill to a preset gradient.


## Syntax

 _expression_ . **PresetGradient**( **_Style_** , **_Variant_** , **_PresetGradientType_** )

 _expression_ Required. A variable that represents a **[FillFormat](fillformat-object-word.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Style_|Required| **MsoGradientStyle**|The gradient style. Can be any  **MsoGradientStyle** constant except **msoGradientFromTitle** which applies only to Microsoft PowerPoint.|
| _Variant_|Required| **Long**|The gradient variant. Can be a value from 1 to 4, corresponding to the four variants on the  **Gradient** tab in the **Fill Effects** dialog box. If Style is **msoGradientFromCenter** , this argument can be either 1 or 2.|
| _PresetGradientType_|Required| **MsoPresetGradientType**|The gradient type.|

## Example

This example adds a rectangle with a preset gradient fill to the active document.


```vb
ActiveDocument.Shapes.AddShape( _ 
 msoShapeRectangle, 90, 90, 140, 80).Fill.PresetGradient _ 
 msoGradientHorizontal, 1, msoGradientBrass
```


## See also


#### Concepts


[FillFormat Object](fillformat-object-word.md)

