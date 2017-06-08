---
title: Column.Shading Property (Word)
keywords: vbawd10.chm156172390
f1_keywords:
- vbawd10.chm156172390
ms.prod: word
api_name:
- Word.Column.Shading
ms.assetid: d85b6720-6be8-6c2d-6e14-7c30c40f83ec
ms.date: 06/08/2017
---


# Column.Shading Property (Word)

Returns a  **Shading** object that refers to the shading formatting for the specified column.


## Syntax

 _expression_ . **Shading**

 _expression_ Required. A variable that represents a **[Column](column-object-word.md)** object.


## Example

This example applies horizontal line texture to the first column in the first table in the active document.


```vb
If ActiveDocument.Tables.Count >= 1 Then 
 With ActiveDocument.Tables(1).Columns(1).Shading 
 .Texture = wdTextureHorizontal 
 End With 
End If
```


## See also


#### Concepts


[Column Object](column-object-word.md)

