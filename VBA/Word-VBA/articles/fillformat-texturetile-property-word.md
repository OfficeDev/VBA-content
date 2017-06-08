---
title: FillFormat.TextureTile Property (Word)
keywords: vbawd10.chm164102264
f1_keywords:
- vbawd10.chm164102264
ms.prod: word
api_name:
- Word.FillFormat.TextureTile
ms.assetid: 670db5f6-8543-2c6e-6aeb-98f240716421
ms.date: 06/08/2017
---


# FillFormat.TextureTile Property (Word)

Returns or sets whether the texture fill is tiled or centered. Read/write.


## Syntax

 _expression_ . **TextureTile**

 _expression_ An expression that returns a **[FillFormat](fillformat-object-word.md)** object.


## Remarks

The value returned by the  **TextureTile** property can be one of the following[MsoTriState](http://msdn.microsoft.com/library/2036cfc9-be7d-e05c-bec7-af05e3c3c515%28Office.15%29.aspx) constants.



|**Constant**|**Description**|
|:-----|:-----|
| **msoFalse**|The texture fill is centered.|
| **msoTrue**|The texture fill is tiled.|
The setting of the  **TextureTile** property corresponds to the setting of the **Tile picture as texture** box under **Tiling Options** on the **Fill** pane of the **Format Picture** dialog box in the Microsoft Word user interface (under **Drawing Tools**, on the  **Format** tab, expand the **Shape Styles** group.)


## See also


#### Concepts


[FillFormat Object](fillformat-object-word.md)

