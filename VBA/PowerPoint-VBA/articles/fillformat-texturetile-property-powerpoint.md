---
title: FillFormat.TextureTile Property (PowerPoint)
keywords: vbapp10.chm552031
f1_keywords:
- vbapp10.chm552031
ms.prod: powerpoint
api_name:
- PowerPoint.FillFormat.TextureTile
ms.assetid: 14d1b329-8d06-b4d6-1ade-aea80f5427ce
ms.date: 06/08/2017
---


# FillFormat.TextureTile Property (PowerPoint)

 Returns or sets whether the texture fill is tiled or centered. Read/write.


## Syntax

 _expression_. **TextureTile**

 _expression_ An expression that returns a **FillFormat** object.


### Return Value

MsoTriState


## Remarks

The value returned by the  **TextureTile** property can be one of these **MsoTriState** constants.



|**Constant**|**Description**|
|:-----|:-----|
|**msoFalse**|The texture fill is centered.|
|**msoTrue**| The texture fill is tiled.|
The setting of the  **TextureTile** property corresponds to the setting of the **Tile picture as texture** box on the **Fill** pane of the **Format Picture** dialog box in the Microsoft PowerPoint user interface (under **Drawing Tools**, on the  **Format Tab**, in the  **Shape Styles** group, click **Format Shape**.)


## See also


#### Concepts


[FillFormat Object](fillformat-object-powerpoint.md)

