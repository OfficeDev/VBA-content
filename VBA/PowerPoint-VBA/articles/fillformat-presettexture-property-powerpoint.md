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

