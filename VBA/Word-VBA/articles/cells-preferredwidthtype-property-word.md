---
title: Cells.PreferredWidthType Property (Word)
keywords: vbawd10.chm155844712
f1_keywords:
- vbawd10.chm155844712
ms.prod: word
api_name:
- Word.Cells.PreferredWidthType
ms.assetid: 65fd3b1d-7048-699b-b549-e2d5265dfe01
ms.date: 06/08/2017
---


# Cells.PreferredWidthType Property (Word)

Returns or sets the preferred unit of measurement to use for the width of the specified cells. Read-only  **WdPreferredWidthType** .


## Syntax

 _expression_ . **PreferredWidthType**

 _expression_ Required. A variable that represents a **[Cells](cells-object-word.md)** collection.


## Example

This example sets Microsoft Word to accept widths as a percentage of window width, and then it sets the width of the first table in the document to 50% of the window width.


```vb
With ActiveDocument.Tables(1) 
 .PreferredWidthType = wdPreferredWidthPercent 
 .PreferredWidth = 50 
End With
```


## See also


#### Concepts


[Cells Collection Object](cells-object-word.md)

