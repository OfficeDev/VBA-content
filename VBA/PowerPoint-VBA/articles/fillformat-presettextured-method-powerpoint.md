---
title: FillFormat.PresetTextured Method (PowerPoint)
keywords: vbapp10.chm552006
f1_keywords:
- vbapp10.chm552006
ms.prod: powerpoint
api_name:
- PowerPoint.FillFormat.PresetTextured
ms.assetid: a025a1d3-a2db-e219-7080-1a29c2fd3f21
ms.date: 06/08/2017
---


# FillFormat.PresetTextured Method (PowerPoint)

Sets the specified fill to a preset texture.


## Syntax

 _expression_. **PresetTextured**( **_PresetTexture_** )

 _expression_ A variable that represents a **FillFormat** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _PresetTexture_|Required|**MsoPresetTexture**|The preset texture.|

## Remarks

The value of the PresetTexturedargument can be one of these  **MsoPresetTexture** constants.


||
|:-----|
|<strong>msoPresetTextureMixed</strong>|
|
<strong>msoTextureBlueTissuePaper</strong>|
|
<strong>msoTextureBouquet</strong>|
|
<strong>msoTextureBrownMarble</strong>|
|
<strong>msoTextureCanvas</strong>|
|
<strong>msoTextureCork</strong>|
|
<strong>msoTextureDenim</strong>|
|
<strong>msoTextureFishFossil</strong>|
|
<strong>msoTextureGranite</strong>|
|
<strong>msoTextureGreenMarble</strong>|
|
<strong>msoTextureMediumWood</strong>|
|
<strong>msoTextureNewsprint</strong>|
|
<strong>msoTextureOak</strong>|
|
<strong>msoTexturePaperBag</strong>|
|
<strong>msoTexturePapyrus</strong>|
|
<strong>msoTextureParchment</strong>|
|
<strong>msoTexturePinkTissuePaper</strong>|
|
<strong>msoTexturePurpleMesh</strong>|
|
<strong>msoTextureRecycledPaper</strong>|
|
<strong>msoTextureSand</strong>|
|
<strong>msoTextureStationery</strong>|
|
<strong>msoTextureWalnut</strong>|
|
<strong>msoTextureWaterDroplets</strong>|
|
<strong>msoTextureWhiteMarble</strong>|
|
<strong>msoTextureWovenMat</strong>|

## Example

This example adds a rectangle with a green-marble textured fill to  `myDocument`.


```vb
Set myDocument = ActivePresentation.Slides(1)

myDocument.Shapes.AddShape(msoShapeCan, 90, 90, 40, 80) _
    .Fill.PresetTextured msoTextureGreenMarble
```


## See also


#### Concepts


[FillFormat Object](fillformat-object-powerpoint.md)

