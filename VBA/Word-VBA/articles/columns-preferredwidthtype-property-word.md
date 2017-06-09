---
title: Columns.PreferredWidthType Property (Word)
keywords: vbawd10.chm155910250
f1_keywords:
- vbawd10.chm155910250
ms.prod: word
api_name:
- Word.Columns.PreferredWidthType
ms.assetid: 2f0a5c0a-177f-5f14-85dc-70e65c020abe
ms.date: 06/08/2017
---


# Columns.PreferredWidthType Property (Word)

Returns or sets the preferred unit of measurement to use for the width of the specified cells, columns, or table. Read/write  **WdPreferredWidthType** .


## Syntax

 _expression_ . **PreferredWidthType**

 _expression_ Required. A variable that represents a **[Columns](columns-object-word.md)** collection.


## Example

This example sets Microsoft Word to accept widths as a percentage of window width, and then it sets the width of all columns in the first table in the active document to 50% of the window width.


```vb
With ActiveDocument.Tables(1).Columns 
 .PreferredWidthType = wdPreferredWidthPercent 
 .PreferredWidth = 50 
End With
```


## See also


#### Concepts


[Columns Collection Object](columns-object-word.md)

