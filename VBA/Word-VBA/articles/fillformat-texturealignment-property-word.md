---
title: FillFormat.TextureAlignment Property (Word)
keywords: vbawd10.chm164102261
f1_keywords:
- vbawd10.chm164102261
ms.prod: word
api_name:
- Word.FillFormat.TextureAlignment
ms.assetid: c28ba99a-8219-996c-676d-ee98d908ab0f
ms.date: 06/08/2017
---


# FillFormat.TextureAlignment Property (Word)

Returns or sets the alignment (the origin of the coordinate grid) for the tiling of the texture fill. Read/write.


## Syntax

 _expression_ . **TextureAlignment**

 _expression_ An expression that returns a **FillFormat** object.


### Return Value

 **MsoTextureAlignment**


## Remarks

The value returned by the  **TextureAlignment** property can be one of the following **MsoTextureAlignment** constants:


- msoTextureTopLeft
    
- msoTextureTop
    
- msoTextureTopRight
    
- msoTextureLeft
    
- msoTextureCenter
    
- msoTextureRight
    
- msoTextureBottomLeft
    
- msoTextureBottom
    
-  msoTextureBottomRight
    
The setting of the  **TextureAlignment** property corresponds to the setting of the **Alignment** box under **Tiling** Options on the **Fill** pane of the ** Format Picture** dialog box in the Microsoft Word user interface (under **Drawing Tools**, on the  **Format Tab,** expand the **Shape Styles** group.)


## See also


#### Concepts


[FillFormat Object](fillformat-object-word.md)

