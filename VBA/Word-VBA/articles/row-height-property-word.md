---
title: Row.Height Property (Word)
keywords: vbawd10.chm156237831
f1_keywords:
- vbawd10.chm156237831
ms.prod: word
api_name:
- Word.Row.Height
ms.assetid: 37586889-891d-5fb4-7f27-d590b92ba77b
ms.date: 06/08/2017
---


# Row.Height Property (Word)

Returns or sets the height (in points) of the specified row in a table. Read/write Single.


## Syntax

 _expression_ . **Height**

 _expression_ A variable that represents a **[Row](row-object-word.md)** object.


## Remarks

 If the **HeightRule** property of the specified row is **wdRowHeightAuto** , **Height** returns **wdUndefined** ; setting the **Height** property sets **HeightRule** to **wdRowHeightAtLeast** .


## Example

This example sets the height of the rows in the first table in the active document to at least 20 points.


```vb
ActiveDocument.Tables(1).Rows.Height = 20
```


## See also


#### Concepts


[Row Object](row-object-word.md)

