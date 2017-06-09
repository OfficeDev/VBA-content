---
title: Rows.Height Property (Word)
keywords: vbawd10.chm155975687
f1_keywords:
- vbawd10.chm155975687
ms.prod: word
api_name:
- Word.Rows.Height
ms.assetid: c111c7e3-0502-118d-035c-be290ea4d83b
ms.date: 06/08/2017
---


# Rows.Height Property (Word)

Returns or sets the height of the specified rows in a table. Read/write Single.


## Syntax

 _expression_ . **Height**

 _expression_ A variable that represents a **[Rows](rows-object-word.md)** collection.


## Remarks

If the  **HeightRule** property of the specified row is **wdRowHeightAuto** , **Height** returns **wdUndefined** ; setting the **Height** property sets **HeightRule** to **wdRowHeightAtLeast** .


## Example

This example sets the height of the rows in the first table in the active document to at least 20 points.


```vb
ActiveDocument.Tables(1).Rows.Height = 20
```


## See also


#### Concepts


[Rows Collection Object](rows-object-word.md)

