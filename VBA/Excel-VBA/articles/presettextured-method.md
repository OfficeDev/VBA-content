---
title: PresetTextured Method
keywords: vbagr10.chm3077629
f1_keywords:
- vbagr10.chm3077629
ms.prod: excel
api_name:
- Excel.PresetTextured
ms.assetid: 4f6abf8c-c09e-6ef8-abb1-0cc643e6458b
ms.date: 06/08/2017
---


# PresetTextured Method

Sets the format of the specified fill to a preset texture.

 _expression_. **PresetTextured**( **_PresetTexture_**)

 _expression_ Required. An expression that returns one of the objects in the Applies To list.

 **PresetTexture**Required 
 **MsoPresetTexture**
. The preset texture for the specified fill.


|MsoPresetTexture can be one of these MsoPresetTexture constants.|
| **msoPresetTextureMixed**|
| **msoTextureBouquet**|
| **msoTextureCanvas**|
| **msoTextureDenim**|
| **msoTextureGranite**|
| **msoTextureMediumWood**|
| **msoTextureOak**|
| **msoTexturePapyrus**|
| **msoTexturePinkTissuePaper**|
| **msoTextureRecycledPaper**|
| **msoTextureStationery**|
| **msoTextureWaterDroplets**|
| **msoTextureWovenMat**|
| **msoTextureBlueTissuePaper**|
| **msoTextureBrownMarble**|
| **msoTextureCork**|
| **msoTextureFishFossil**|
| **msoTextureGreenMarble**|
| **msoTextureNewsprint**|
| **msoTexturePaperBag**|
| **msoTextureParchment**|
| **msoTexturePurpleMesh**|
| **msoTextureSand**|
| **msoTextureWalnut**|
| **msoTextureWhiteMarble**|

## Example

This example changes the chart's textured fill format from oak to walnut.


```vb
With myChart.ChartArea.Fill 
 If .Type = msoFillTextured Then 
 If .TextureType = msoTexturePreset Then 
 If .PresetTexture = msoTextureOak Then 
 .PresetTextured msoTextureWalnut 
 End If 
 End If 
 End If 
End With
```


