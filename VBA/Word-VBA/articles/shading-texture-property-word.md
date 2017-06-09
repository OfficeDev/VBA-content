---
title: Shading.Texture Property (Word)
keywords: vbawd10.chm154796035
f1_keywords:
- vbawd10.chm154796035
ms.prod: word
api_name:
- Word.Shading.Texture
ms.assetid: 97fac431-4e0a-fd92-9845-47ee99196a78
ms.date: 06/08/2017
---


# Shading.Texture Property (Word)

Returns or sets the shading texture for the specified object. Read/write  **WdTextureIndex** .


## Syntax

 _expression_ . **Texture**

 _expression_ Required. A variable that represents a **[Shading](shading-object-word.md)** object.


## Example

This example sets a range that references the first paragraph in the active document and then applies a grid texture to that range.


```vb
Set myRange = ActiveDocument.Paragraphs(1).Range 
myRange.Shading.Texture = wdTextureCross
```

This example adds a table at the insertion point and then applies a vertical line texture to the first row in the table.




```vb
Selection.Collapse Direction:=wdCollapseStart 
Set myTable = ActiveDocument.Tables.Add(Range:=Selection.Range, _ 
 NumRows:=2, NumColumns:=2) 
myTable.Rows(1).Shading.Texture = wdTextureVertical
```

This example applies 10 percent shading to the first word in the active document.




```vb
ActiveDocument.Words(1).Shading.Texture = wdTexture10Percent
```


## See also


#### Concepts


[Shading Object](shading-object-word.md)

