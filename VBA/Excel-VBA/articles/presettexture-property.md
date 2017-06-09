---
title: PresetTexture Property
keywords: vbagr10.chm67162
f1_keywords:
- vbagr10.chm67162
ms.prod: excel
api_name:
- Excel.PresetTexture
ms.assetid: 5b471290-66f4-3504-096b-70265db88b93
ms.date: 06/08/2017
---


# PresetTexture Property

Returns the preset texture for the specified fill. Read-only MsoPresetTexture .



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

 _expression_. **PresetTexture**

 _expression_ Required. An expression that returns one of the objects in the Applies To list.
This property is read-only. Use the  **PresetTextured** method to set the preset texture for the fill.

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


