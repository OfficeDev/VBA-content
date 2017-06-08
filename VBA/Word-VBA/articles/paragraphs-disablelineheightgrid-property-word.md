---
title: Paragraphs.DisableLineHeightGrid Property (Word)
keywords: vbawd10.chm156762237
f1_keywords:
- vbawd10.chm156762237
ms.prod: word
api_name:
- Word.Paragraphs.DisableLineHeightGrid
ms.assetid: 287370a2-bf08-0104-ec28-ba9e934a8848
ms.date: 06/08/2017
---


# Paragraphs.DisableLineHeightGrid Property (Word)

 **True** if Microsoft Word aligns characters in the specified paragraphs to the line grid when a set number of lines per page is specified. Returns **wdUndefined** if the **DisableLineHeightGrid** property is set to **True** for only some of the specified paragraphs. Read/write **Long** .


## Syntax

 _expression_ . **DisableLineHeightGrid**

 _expression_ A variable that represents a **[Paragraphs](paragraphs-object-word.md)** collection.


## Example

This example sets Word to align characters in the selected paragraphs to the line grid if you've specified a set number of lines per page.


```vb
With Selection.ParagraphFormat 
 .DisableLineHeightGrid = True 
End With
```


## See also


#### Concepts


[Paragraphs Collection Object](paragraphs-object-word.md)

