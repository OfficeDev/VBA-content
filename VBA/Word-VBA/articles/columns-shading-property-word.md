---
title: Columns.Shading Property (Word)
keywords: vbawd10.chm155910247
f1_keywords:
- vbawd10.chm155910247
ms.prod: word
api_name:
- Word.Columns.Shading
ms.assetid: 8dd27658-7208-86ae-09b1-bf4f89280402
ms.date: 06/08/2017
---


# Columns.Shading Property (Word)

Returns a  **Shading** object that refers to the shading formatting for the specified table columns.


## Syntax

 _expression_ . **Shading**

 _expression_ Required. A variable that represents a **[Columns](columns-object-word.md)** collection.


## Example

This example applies horizontal line texture to all columns in the first table in the active document.


```vb
If ActiveDocument.Tables.Count >= 1 Then 
 With ActiveDocument.Tables(1).Columns.Shading 
 .Texture = wdTextureDiagonalDown 
 End With 
End If
```


## See also


#### Concepts


[Columns Collection Object](columns-object-word.md)

