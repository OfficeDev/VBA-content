---
title: Table.PreferredWidthType Property (Word)
keywords: vbawd10.chm156303472
f1_keywords:
- vbawd10.chm156303472
ms.prod: word
api_name:
- Word.Table.PreferredWidthType
ms.assetid: 92954057-5ecd-3d43-c547-e1e1a6c83904
ms.date: 06/08/2017
---


# Table.PreferredWidthType Property (Word)

Returns or sets the preferred unit of measurement to use for the width of the specified table. Read/write  **[WdPreferredWidthType](wdpreferredwidthtype-enumeration-word.md)** .


## Syntax

 _expression_ . **PreferredWidthType**

 _expression_ Required. A variable that represents a **[Table](table-object-word.md)** object.


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


[Table Object](table-object-word.md)

