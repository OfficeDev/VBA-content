---
title: FillFormat.PresetTexture Property (PowerPoint)
keywords: vbapp10.chm552019
f1_keywords:
- vbapp10.chm552019
ms.prod: powerpoint
api_name:
- PowerPoint.FillFormat.PresetTexture
ms.assetid: 684d39f9-53d8-4f69-a6ae-c447253ae3a7
ms.date: 06/08/2017
---


# FillFormat.PresetTexture Property (PowerPoint)

Returns the preset texture for the specified fill. Read-only.


## Syntax

 _expression_. **PresetTexture**

 _expression_ A variable that represents a **FillFormat** object.


### Return Value

MsoPresetTexture


## Remarks

The value of the  **PresetTexture** property can be one of these **MsoPresetTexture** constants.


||
|:-----|
|**msoPresetTextureMixed**|
|**msoTextureBlueTissuePaper**|
|**msoTextureBouquet**|
|**msoTextureBrownMarble**|
|**msoTextureCanvas**|
|**msoTextureCork**|
|**msoTextureDenim**|
|**msoTextureFishFossil**|
|**msoTextureGranite**|
|**msoTextureGreenMarble**|
|**msoTextureMediumWood**|
|**msoTextureNewsprint**|
|**msoTextureOak**|
|**msoTexturePaperBag**|
|**msoTexturePapyrus**|
|**msoTextureParchment**|
|**msoTexturePinkTissuePaper**|
|**msoTexturePurpleMesh**|
|**msoTextureRecycledPaper**|
|**msoTextureSand**|
|**msoTextureStationery**|
|**msoTextureWalnut**|
|**msoTextureWaterDroplets**|
|**msoTextureWhiteMarble**|
|**msoTextureWovenMat**|

## Example

This example adds a rectangle to the  `myDocument` and sets its preset texture to match that of shape two. For the example to work, shape two must have a preset textured fill.


```vb
Set myDocument = ActivePresentation.Slides(1)

With myDocument.Shapes
    presetTexture2 = .Item(2).Fill.PresetTexture
    .AddShape(msoShapeRectangle, 100, 0, 40, 80).Fill _
        .PresetTextured presetTexture2
End With
```


## See also


#### Concepts


[FillFormat Object](fillformat-object-powerpoint.md)

