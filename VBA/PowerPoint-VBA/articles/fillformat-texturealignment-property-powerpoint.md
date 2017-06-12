---
title: FillFormat.TextureAlignment Property (PowerPoint)
keywords: vbapp10.chm552028
f1_keywords:
- vbapp10.chm552028
ms.prod: powerpoint
api_name:
- PowerPoint.FillFormat.TextureAlignment
ms.assetid: e26ca83c-7dc1-4c7b-52a4-3a30669079ea
ms.date: 06/08/2017
---


# FillFormat.TextureAlignment Property (PowerPoint)

Returns or sets the alignment (the origin of the coordinate grid) for the tiling of the texture fill. Read/write.


## Syntax

 _expression_. **TextureAlignment**

 _expression_ An expression that returns a **FillFormat** object.


### Return Value

MsoTextureAlignment


## Remarks

The value returned by the  **TextureAlignment** property can be one of these **MsoTextureAlignment** constants.


||
|:-----|
|**msoTextureTopLeft**|
|**msoTextureTop**|
|**msoTextureTopRight**|
|**msoTextureLeft**|
|**msoTextureCenter**|
|**msoTextureRight**|
|**msoTextureBottomLeft**|
|**msoTextureBottom**|
|**msoTextureBottomRight**|
The setting of the  **TextureAlignment** property corresponds to the setting of the **Alignment** box under **Tiling Options** on the **Fill** pane of the **Format Picture** dialog box in the Microsoft PowerPoint user interface (under **Drawing Tools**, on the  **Format Tab**, in the  **Shape Styles** group, click **Format Shape**.)


## See also


#### Concepts


[FillFormat Object](fillformat-object-powerpoint.md)

