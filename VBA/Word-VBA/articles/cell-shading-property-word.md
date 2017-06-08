---
title: Cell.Shading Property (Word)
keywords: vbawd10.chm156106857
f1_keywords:
- vbawd10.chm156106857
ms.prod: word
api_name:
- Word.Cell.Shading
ms.assetid: ab2f5789-ba6e-fa8a-d0a9-4c8b7922aa92
ms.date: 06/08/2017
---


# Cell.Shading Property (Word)

Returns a  **[Shading](shading-object-word.md)** object that refers to the shading formatting for the specified object.


## Syntax

 _expression_ . **Shading**

 _expression_ A variable that represents a **[Cell](cell-object-word.md)** object.


## Example

This example applies horizontal line texture to the first cell in the first row in first table.


```vb
If ActiveDocument.Tables.Count >= 1 Then 
 With ActiveDocument.Tables(1).Rows(1).Cells(1).Shading 
 .Texture = wdTextureHorizontal 
 End With 
End If
```


## See also


#### Concepts


[Cell Object](cell-object-word.md)

