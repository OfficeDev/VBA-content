---
title: Cells.Height Property (Word)
keywords: vbawd10.chm155844615
f1_keywords:
- vbawd10.chm155844615
ms.prod: word
api_name:
- Word.Cells.Height
ms.assetid: 54577b7c-2b68-1054-958a-49dd0fb76978
ms.date: 06/08/2017
---


# Cells.Height Property (Word)

Returns or sets the height of the specified table cells. Read/write  **Single** .


## Syntax

 _expression_ . **Height**

 _expression_ An expression that returns a **[Cells](cells-object-word.md)** collection.


## Remarks

If the  **HeightRule** property of the specified row is **wdRowHeightAuto** , **Height** returns **wdUndefined** ; setting the **Height** property sets **HeightRule** to **wdRowHeightAtLeast** .


## See also


#### Concepts


[Cells Collection Object](cells-object-word.md)

