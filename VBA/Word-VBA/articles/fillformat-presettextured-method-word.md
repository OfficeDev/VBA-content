---
title: FillFormat.PresetTextured Method (Word)
keywords: vbawd10.chm164102158
f1_keywords:
- vbawd10.chm164102158
ms.prod: word
api_name:
- Word.FillFormat.PresetTextured
ms.assetid: 9a4aac9d-6349-7947-bc4a-1b0bb64a848b
ms.date: 06/08/2017
---


# FillFormat.PresetTextured Method (Word)

Sets the specified fill to a preset texture.


## Syntax

 _expression_ . **PresetTextured**( **_PresetTexture_** )

 _expression_ Required. A variable that represents a **[FillFormat](fillformat-object-word.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _PresetTexture_|Required| **MsoPresetTexture**|The preset texture.|

## Example

This example adds a rectangle with a green-marble textured fill to the active document.


```vb
ActiveDocument.Shapes.AddShape(msoShapeCan, 90, 90, 40, 80) _ 
 .Fill.PresetTextured msoTextureGreenMarble
```


## See also


#### Concepts


[FillFormat Object](fillformat-object-word.md)

