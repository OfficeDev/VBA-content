---
title: Table.PreferredWidth Property (Word)
keywords: vbawd10.chm156303471
f1_keywords:
- vbawd10.chm156303471
ms.prod: word
api_name:
- Word.Table.PreferredWidth
ms.assetid: 15c3d169-9c61-fb70-3cc6-15f385bab8c0
ms.date: 06/08/2017
---


# Table.PreferredWidth Property (Word)

Returns or sets the preferred width (in points or as a percentage of the window width) for the specified table. Read/write  **Single** .


## Syntax

 _expression_ . **PreferredWidth**

 _expression_ An expression that represents a **[Table](table-object-word.md)** object.


## Remarks

If the  **[PreferredWidthType](table-preferredwidthtype-property-word.md)** property is set to **wdPreferredWidthPoints** , the **PreferredWidth** property returns or sets the width in points. If the **PreferredWidthType** property is set to **wdPreferredWidthPercent** , the **PreferredWidth** property returns or sets the width as a percentage of the window width.


## Example

This example sets Microsoft Word to accept preferred widths as a percentage of window width, and then sets the preferred width of the first table in the document to 50% of the window width.


```vb
With ActiveDocument.Tables(1) 
 .PreferredWidthType = wdPreferredWidthPercent 
 .PreferredWidth = 50 
End With
```


## See also


#### Concepts


[Table Object](table-object-word.md)

