---
title: FillFormat.PresetTextured Method (Publisher)
keywords: vbapb10.chm2359316
f1_keywords:
- vbapb10.chm2359316
ms.prod: publisher
api_name:
- Publisher.FillFormat.PresetTextured
ms.assetid: 971eac34-4e29-c898-93c8-9e71bd92238d
ms.date: 06/08/2017
---


# FillFormat.PresetTextured Method (Publisher)

Sets the specified fill to a preset texture.


## Syntax

 _expression_. **PresetTextured**( **_PresetTexture_**)

 _expression_A variable that represents a  **FillFormat** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
|PresetTexture|Required| **MsoPresetTexture**|The preset texture.|

## Remarks

The PresetTexture parameter can be one of the following  **MsoPresetTexture** constants declared in the Microsoft Office type library.



| **msoTextureBlueTissuePaper**|
| **msoTextureBouquet**|
| **msoTextureBrownMarble**|
| **msoTextureCanvas**|
| **msoTextureCork**|
| **msoTextureDenim**|
| **msoTextureFishFossil**|
| **msoTextureGranite**|
| **msoTextureGreenMarble**|
| **msoTextureMediumWood**|
| **msoTextureNewsprint**|
| **msoTextureOak**|
| **msoTexturePaperBag**|
| **msoTexturePapyrus**|
| **msoTextureParchment**|
| **msoTexturePinkTissuePaper**|
| **msoTexturePurpleMesh**|
| **msoTextureRecycledPaper**|
| **msoTextureSand**|
| **msoTextureStationery**|
| **msoTextureWalnut**|
| **msoTextureWaterDroplets**|
| **msoTextureWhiteMarble**|
| **msoTextureWovenMat**|

## Example

This example adds a rectangle with a green-marble textured fill to the active publication.


```vb
ActiveDocument.Pages(1).Shapes _ 
 .AddShape(Type:=msoShapeCan, _ 
 Left:=90, Top:=90, Width:=40, Height:=80) _ 
 .Fill.PresetTextured _ 
 PresetTexture:=msoTextureGreenMarble 

```


